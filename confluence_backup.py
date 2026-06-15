"""
Confluence 페이지 + 하위 페이지 로컬 백업 스크립트
- 출력 포맷: Markdown (본문) + JSON (메타데이터)
- 이미지: 원본 저장 + Claude Vision으로 텍스트 설명 인라인 삽입
- AI 분석·활용에 최적화된 구조로 저장

사용법:
    pip install requests markdownify anthropic

    python confluence_backup.py \
        --url https://your-domain.atlassian.net \
        --user your@email.com \
        --token YOUR_API_TOKEN \
        --page-id 3121812261 \
        --output ./backup \
        [--anthropic-key YOUR_ANTHROPIC_KEY]  # 이미지 Vision 처리 시 필요
"""

import os
import re
import json
import time
import base64
import argparse
import logging
from pathlib import Path
from datetime import datetime
from typing import Optional

import requests
from markdownify import markdownify as md

# ── 로깅 설정 ────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


# ═══════════════════════════════════════════════════════════════════════════════
# 설정
# ═══════════════════════════════════════════════════════════════════════════════

SUPPORTED_IMAGE_TYPES = {"image/png", "image/jpeg", "image/jpg", "image/gif", "image/webp"}
MAX_VISION_RETRIES = 3
REQUEST_DELAY = 0.3          # Confluence API 호출 간격 (초)


# ═══════════════════════════════════════════════════════════════════════════════
# Confluence API 클라이언트
# ═══════════════════════════════════════════════════════════════════════════════

class ConfluenceClient:
    def __init__(self, base_url: str, email: str, token: str, auth_type: str = "basic"):
        base = base_url.rstrip("/")

        self.session = requests.Session()
        self.session.headers.update({"Accept": "application/json"})

        # Cloud  : Basic Auth (email + API token)
        # Server : Bearer Token (PAT, email 불필요)
        if auth_type == "bearer":
            self.session.headers["Authorization"] = f"Bearer {token}"
            log.info("인증 방식: Bearer Token (Server/Data Center)")
        else:
            self.session.auth = (email, token)
            log.info("인증 방식: Basic Auth (Cloud)")

        # API base URL 자동 감지
        # 페이지 URL 패턴 예시:
        #   http://mydomain.com/main/spaces/.../pages/ID  → context_path = /main
        #   https://myco.atlassian.net/wiki/spaces/...    → context_path = /wiki
        #   https://myco.atlassian.net                    → context_path = /wiki (Cloud 기본)
        self.base_url, self.api_base = self._resolve_api_base(base)
        log.info(f"API base URL: {self.api_base}")

    def _resolve_api_base(self, base: str) -> tuple[str, str]:
        """
        URL에서 context path를 추출하여 API base URL 결정.
        /spaces/ 또는 /pages/ 앞의 경로를 context path로 사용.
        """
        from urllib.parse import urlparse

        parsed = urlparse(base)
        path = parsed.path.rstrip("/")
        origin = f"{parsed.scheme}://{parsed.netloc}"

        # /spaces/ 또는 /pages/ 가 포함된 경우 그 앞까지를 context path로 사용
        for marker in ("/spaces/", "/pages/"):
            if marker in path:
                context = path.split(marker)[0]  # 예: /main
                api_base = f"{origin}{context}/rest/api"
                return origin, api_base

        # /wiki 로 끝나는 경우
        if path.endswith("/wiki"):
            context = path  # 예: /wiki
            api_base = f"{origin}{context}/rest/api"
            return origin, api_base

        # context path가 있는 경우 (예: /main)
        if path and path != "/":
            api_base = f"{origin}{path}/rest/api"
            return origin, api_base

        # Cloud 기본: 도메인만 있는 경우
        api_base = f"{origin}/wiki/rest/api"
        return origin, api_base

    def _get(self, path: str, params: dict = None) -> dict:
        url = f"{self.api_base}{path}"
        log.debug(f"GET {url}  params={params}")
        resp = self.session.get(url, params=params, timeout=30)
        if not resp.ok:
            log.error(f"HTTP {resp.status_code} {resp.reason} — {url}")
            log.error(f"응답 본문: {resp.text[:300]}")
        resp.raise_for_status()
        time.sleep(REQUEST_DELAY)
        return resp.json()

    def get_page(self, page_id: str) -> dict:
        """페이지 본문 + 메타데이터 조회"""
        return self._get(
            f"/content/{page_id}",
            params={
                "expand": (
                    "body.view,"
                    "metadata.labels,"
                    "version,"
                    "space,"
                    "ancestors,"
                    "children.attachment"
                )
            },
        )

    def get_children(self, page_id: str) -> list[dict]:
        """직접 하위 페이지 목록 조회 (페이지네이션 처리)"""
        children = []
        start = 0
        limit = 50
        while True:
            data = self._get(
                f"/content/{page_id}/child/page",
                params={"limit": limit, "start": start, "expand": "version"},
            )
            children.extend(data.get("results", []))
            if len(data.get("results", [])) < limit:
                break
            start += limit
        return children

    def download_attachment(self, download_path: str) -> bytes:
        """첨부파일 바이너리 다운로드"""
        # download_path 가 /wiki/... 또는 /main/... 형태로 오므로 origin에 바로 붙임
        url = f"{self.base_url}{download_path}"
        resp = self.session.get(url, timeout=60)
        resp.raise_for_status()
        return resp.content


# ═══════════════════════════════════════════════════════════════════════════════
# Vision 처리 (Claude API)
# ═══════════════════════════════════════════════════════════════════════════════

class VisionProcessor:
    def __init__(self, api_key: Optional[str]):
        self.enabled = bool(api_key)
        if self.enabled:
            try:
                import anthropic
                self.client = anthropic.Anthropic(api_key=api_key)
            except ImportError:
                log.warning("anthropic 패키지 미설치. Vision 처리를 건너뜁니다.")
                self.enabled = False

    def describe(self, image_bytes: bytes, media_type: str, filename: str) -> Optional[str]:
        """이미지를 텍스트 설명으로 변환"""
        if not self.enabled:
            return None

        b64 = base64.standard_b64encode(image_bytes).decode()
        prompt = (
            "이 이미지를 AI 검색·분석 시스템이 활용할 수 있도록 상세히 설명해주세요.\n"
            "포함 항목:\n"
            "- 이미지 유형 (다이어그램/차트/스크린샷/사진 등)\n"
            "- 포함된 모든 텍스트, 수치, 레이블\n"
            "- 구조나 흐름 (화살표, 계층, 관계 등)\n"
            "- 핵심 정보 요약\n"
            "설명만 출력하고, 전문(preamble)은 생략하세요."
        )

        for attempt in range(MAX_VISION_RETRIES):
            try:
                resp = self.client.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=800,
                    messages=[{
                        "role": "user",
                        "content": [
                            {
                                "type": "image",
                                "source": {
                                    "type": "base64",
                                    "media_type": media_type,
                                    "data": b64,
                                },
                            },
                            {"type": "text", "text": prompt},
                        ],
                    }],
                )
                return resp.content[0].text.strip()
            except Exception as e:
                log.warning(f"Vision 처리 실패 ({filename}, 시도 {attempt+1}/{MAX_VISION_RETRIES}): {e}")
                time.sleep(2 ** attempt)

        return None


# ═══════════════════════════════════════════════════════════════════════════════
# 파일명 정규화
# ═══════════════════════════════════════════════════════════════════════════════

def safe_filename(name: str, max_len: int = 80) -> str:
    """파일시스템에 안전한 파일명 생성"""
    name = re.sub(r'[\\/*?:"<>|]', "_", name)
    name = re.sub(r"\s+", "_", name).strip("._")
    return name[:max_len] or "untitled"


# ═══════════════════════════════════════════════════════════════════════════════
# 백업 코어
# ═══════════════════════════════════════════════════════════════════════════════

class ConfluenceBackup:
    def __init__(
        self,
        client: ConfluenceClient,
        vision: VisionProcessor,
        output_dir: Path,
    ):
        self.client = client
        self.vision = vision
        self.output_dir = output_dir
        self.pages_dir = output_dir / "pages"
        self.attachments_dir = output_dir / "attachments"
        self.metadata_dir = output_dir / "metadata"
        self.index: list[dict] = []

        for d in (self.pages_dir, self.attachments_dir, self.metadata_dir):
            d.mkdir(parents=True, exist_ok=True)

    # ── 단일 페이지 처리 ────────────────────────────────────────────────────

    def _process_attachments(self, page_id: str, attachments: list[dict]) -> dict[str, str]:
        """
        이미지 첨부파일 다운로드 + Vision 설명 생성
        반환: {파일명: Vision 설명 텍스트}
        """
        att_page_dir = self.attachments_dir / page_id
        att_page_dir.mkdir(exist_ok=True)

        descriptions: dict[str, str] = {}

        for att in attachments:
            media_type = att.get("metadata", {}).get("mediaType", "")
            if media_type not in SUPPORTED_IMAGE_TYPES:
                continue  # 이미지가 아닌 첨부파일은 건너뜀

            title = att.get("title", "unknown")
            download_link = att.get("_links", {}).get("download", "")

            # 1. 다운로드
            try:
                img_bytes = self.client.download_attachment(download_link)
            except Exception as e:
                log.warning(f"  첨부파일 다운로드 실패 ({title}): {e}")
                continue

            # 2. 로컬 저장
            save_path = att_page_dir / safe_filename(title)
            save_path.write_bytes(img_bytes)
            log.info(f"  🖼️  이미지 저장: {save_path.name}")

            # 3. Vision 설명
            desc = self.vision.describe(img_bytes, media_type, title)
            if desc:
                descriptions[title] = desc
                log.info(f"  ✨ Vision 설명 생성: {title}")
            else:
                descriptions[title] = f"[이미지: {title}]"

        return descriptions

    def _html_to_markdown(self, html: str, image_descriptions: dict[str, str]) -> str:
        """
        HTML → Markdown 변환 후, 이미지 참조 위치에 Vision 설명 인라인 삽입
        """
        # Confluence 전용 매크로 태그 간략 제거
        html = re.sub(r"<ac:[^>]+>.*?</ac:[^>]+>", "", html, flags=re.DOTALL)
        html = re.sub(r"<ri:[^>]+/>", "", html)

        markdown = md(html, heading_style="ATX", bullets="-", newline_style="backslash")

        # 이미지 마크다운 패턴: ![alt](url) → 설명 블록 삽입
        def replace_image(match):
            alt = match.group(1)
            # alt 또는 파일명으로 설명 찾기
            desc = image_descriptions.get(alt) or next(
                (v for k, v in image_descriptions.items() if alt in k or k in alt), None
            )
            if desc:
                return f"{match.group(0)}\n\n> 🖼️ **이미지 설명**: {desc}\n"
            return match.group(0)

        markdown = re.sub(r"!\[([^\]]*)\]\([^)]*\)", replace_image, markdown)

        # 연속 빈 줄 정리
        markdown = re.sub(r"\n{3,}", "\n\n", markdown)
        return markdown.strip()

    def backup_page(self, page_id: str, depth: int = 0) -> Optional[dict]:
        """단일 페이지 백업 (재귀 호출용)"""
        indent = "  " * depth
        try:
            data = self.client.get_page(page_id)
        except Exception as e:
            log.error(f"{indent}페이지 조회 실패 (ID: {page_id}): {e}")
            return None

        title = data.get("title", "untitled")
        log.info(f"{indent}📄 백업 중: {title} (ID: {page_id})")

        # 첨부 이미지 처리
        raw_attachments = data.get("children", {}).get("attachment", {}).get("results", [])
        image_descriptions = self._process_attachments(page_id, raw_attachments)

        # 본문 HTML → Markdown
        html_body = data.get("body", {}).get("view", {}).get("value", "")
        markdown_body = self._html_to_markdown(html_body, image_descriptions)

        # 메타데이터 구성
        ancestors = [
            {"id": a["id"], "title": a["title"]}
            for a in data.get("ancestors", [])
        ]
        labels = [
            lb["name"]
            for lb in data.get("metadata", {}).get("labels", {}).get("results", [])
        ]
        meta = {
            "id": page_id,
            "title": title,
            "space": data.get("space", {}).get("key", ""),
            "version": data.get("version", {}).get("number"),
            "created_by": data.get("version", {}).get("by", {}).get("displayName", ""),
            "last_modified": data.get("version", {}).get("when", ""),
            "labels": labels,
            "ancestors": ancestors,
            "depth": depth,
            "has_images": bool(image_descriptions),
            "image_count": len(image_descriptions),
            "backed_up_at": datetime.now().isoformat(),
        }

        # ── 파일 저장 ────────────────────────────────────────────────────────

        file_stem = f"{page_id}_{safe_filename(title)}"

        # Markdown 본문
        md_path = self.pages_dir / f"{file_stem}.md"
        md_content = f"# {title}\n\n{markdown_body}"
        md_path.write_text(md_content, encoding="utf-8")

        # JSON 메타데이터
        json_path = self.metadata_dir / f"{page_id}.json"
        json_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")

        # 인덱스 등록
        index_entry = {**meta, "md_file": str(md_path.relative_to(self.output_dir))}
        self.index.append(index_entry)

        return meta

    # ── 재귀 백업 ───────────────────────────────────────────────────────────

    def backup_tree(self, page_id: str, depth: int = 0):
        """루트 페이지 + 모든 하위 페이지 재귀 백업"""
        self.backup_page(page_id, depth)

        try:
            children = self.client.get_children(page_id)
        except Exception as e:
            log.warning(f"하위 페이지 조회 실패 (ID: {page_id}): {e}")
            return

        for child in children:
            self.backup_tree(child["id"], depth + 1)

    # ── 인덱스 저장 ─────────────────────────────────────────────────────────

    def save_index(self):
        """전체 백업 인덱스 JSON 저장"""
        index_path = self.output_dir / "index.json"
        index_data = {
            "backed_up_at": datetime.now().isoformat(),
            "total_pages": len(self.index),
            "pages": self.index,
        }
        index_path.write_text(json.dumps(index_data, ensure_ascii=False, indent=2), encoding="utf-8")
        log.info(f"\n📋 인덱스 저장: {index_path}")


# ═══════════════════════════════════════════════════════════════════════════════
# 진입점
# ═══════════════════════════════════════════════════════════════════════════════

def parse_args():
    """
    - --url, --user, --page-id : CLI 필수 인자
    - ATLASSIAN_CONFLUENCE_TOKEN          : 환경변수 필수 (CLI 인자 없음)
    - ANTHROPIC_API_KEY         : 환경변수 선택 (Vision 처리 시)
    """
    parser = argparse.ArgumentParser(
        description="Confluence 페이지 트리를 AI 활용에 최적화된 형태로 로컬 백업",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
사전 환경변수 설정 (필수):
  export ATLASSIAN_CONFLUENCE_TOKEN=YOUR_API_TOKEN

선택 환경변수 (Vision 처리 시):
  export ANTHROPIC_API_KEY=sk-ant-...

실행 예시:
  python confluence_backup.py \\
      --url https://mycompany.atlassian.net \\
      --user your@email.com \\
      --page-id 123456789 \\
      --output ./backup
        """,
    )
    parser.add_argument("--url",           required=True, help="Confluence 도메인 (예: https://mycompany.atlassian.net)")
    parser.add_argument("--user",          required=True, help="Atlassian 계정 이메일")
    parser.add_argument("--page-id",       required=True, help="백업 시작 페이지 ID")
    parser.add_argument("--output",        default="./confluence_backup", help="백업 저장 경로 (기본: ./confluence_backup)")
    parser.add_argument("--anthropic-key", default=None,  help="Claude Vision API 키 (생략 시 ANTHROPIC_API_KEY 환경변수 사용)")
    parser.add_argument("--auth-type",     default="basic", choices=["basic", "bearer"],
                        help="인증 방식: basic=Cloud(기본), bearer=Server/Data Center")
    parser.add_argument("--debug",         action="store_true", help="디버그 로그 출력 (URL·HTTP 요청 상세)")
    return parser.parse_args()


def resolve_config(args) -> dict:
    """CLI 인자 + 환경변수(토큰)를 합쳐 최종 설정값 반환"""
    token = os.environ.get("ATLASSIAN_CONFLUENCE_TOKEN")
    if not token:
        raise SystemExit(
            "\n❌ ATLASSIAN_CONFLUENCE_TOKEN 환경변수가 설정되지 않았습니다.\n"
            "   export ATLASSIAN_CONFLUENCE_TOKEN=YOUR_API_TOKEN\n"
        )

    return {
        "url":           args.url,
        "user":          args.user,
        "token":         token,
        "page_id":       args.page_id,
        "output":        args.output,
        "auth_type":     args.auth_type,
        "anthropic_key": args.anthropic_key or os.environ.get("ANTHROPIC_API_KEY"),
    }


def main():
    args = parse_args()
    cfg  = resolve_config(args)

    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)

    output_dir = Path(cfg["output"])
    output_dir.mkdir(parents=True, exist_ok=True)

    log.info("=" * 60)
    log.info("Confluence 백업 시작")
    log.info(f"  대상 페이지 ID : {cfg['page_id']}")
    log.info(f"  저장 경로      : {output_dir.resolve()}")
    log.info(f"  Vision 처리    : {'활성화' if cfg['anthropic_key'] else '비활성화'}")
    log.info("=" * 60)

    client = ConfluenceClient(cfg["url"], cfg["user"], cfg["token"], cfg["auth_type"])
    vision = VisionProcessor(cfg["anthropic_key"])
    backup = ConfluenceBackup(client, vision, output_dir)

    start = time.time()
    backup.backup_tree(cfg["page_id"])
    backup.save_index()
    elapsed = time.time() - start

    log.info("=" * 60)
    log.info(f"✅ 백업 완료: {len(backup.index)}개 페이지 | 소요 시간: {elapsed:.1f}초")
    log.info(f"📁 저장 위치: {output_dir.resolve()}")
    log.info("=" * 60)

    # 저장 구조 출력
    print("\n📂 백업 디렉토리 구조:")
    print(f"  {output_dir}/")
    print(f"  ├── index.json                 ← 전체 페이지 목록 + 검색용")
    print(f"  ├── pages/                     ← Markdown 본문 (이미지 설명 포함)")
    print(f"  ├── attachments/               ← 원본 이미지 파일")
    print(f"  └── metadata/                  ← 페이지별 JSON 메타데이터")


if __name__ == "__main__":
    main()