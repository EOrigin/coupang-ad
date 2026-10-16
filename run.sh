#!/usr/bin/env bash

# 스크립트 위치로 이동
cd "$(dirname "$0")"

# 가상환경 활성화
if [ -d ".venv" ]; then
    source .venv/bin/activate
fi

# 크롤러 실행 (전달받은 모든 인자 전달)
python coupang_ad_crawler.py "$@"
