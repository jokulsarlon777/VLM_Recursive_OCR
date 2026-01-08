# Quick Start Guide

## 대용량 데이터셋 (1만개 이상) - 권장 방법

### 1. 설치
```bash
pip install -r requirements.txt
copy .env.example .env
# .env 파일을 열어 OPENAI_API_KEY 입력
```

### 2. 데이터 준비
```bash
# PowerPoint 파일을 data/ 디렉토리에 배치
```

### 3. Step 1: 이미지 변환 실행
```bash
python step1_convert.py
```

**결과:**
- `temp/` 디렉토리에 모든 슬라이드 이미지 생성
- `output/step1_metadata.json` 메타데이터 저장
- **예상 시간**: 1만개 슬라이드 기준 약 3-5시간

### 4. Step 2: VLM 분석 실행
```bash
# 기본 설정 (5 workers)
python step2_analyze.py

# 더 빠르게 (10 workers - 권장)
python step2_analyze.py --workers 10

# 매우 빠르게 (20 workers - API rate limit 주의)
python step2_analyze.py --workers 20
```

**결과:**
- `output/{filename}_d{depth}_analysis.json` 개별 분석 결과
- `output/processing_summary.json` 전체 요약
- **예상 시간**: 1만개 슬라이드 기준 약 2-4시간 (worker 10개)

---

## 소규모 데이터셋 - 간단한 방법

### 한번에 실행
```bash
python main.py
```

### Worker 수 조정
```bash
python main.py --workers 10
```

---

## 재실행 시나리오

### Step 2만 다시 실행 (이미지 재변환 불필요)
```bash
# VLM 분석만 다시 실행
python step2_analyze.py --workers 5

# 또는
python main.py --skip-step1 --workers 10
```

### Step 1만 다시 실행 (VLM 분석 건너뛰기)
```bash
python step1_convert.py

# 또는
python main.py --skip-step2
```

---

## 문제 해결

### "Metadata file not found" 에러
```bash
# Step 2를 먼저 실행했을 때 발생
# 해결: Step 1을 먼저 실행하세요
python step1_convert.py
```

### API Rate Limit 에러
```bash
# Worker 수를 줄이세요
python step2_analyze.py --workers 3
```

### 중간에 중단됨
```bash
# 다시 실행하면 됩니다
# Step 1: 이미 처리된 파일은 자동으로 건너뜀
# Step 2: 메타데이터가 있으면 언제든 재실행 가능
python step2_analyze.py
```

---

## 비용 절약 팁

1. **Step 1 먼저 완료**: 이미지 변환은 무료, VLM 분석만 비용 발생
2. **소량 테스트**: 샘플 데이터로 먼저 테스트
3. **Worker 수 조정**: 적은 수로 시작해서 점진적으로 증가
4. **Step 2만 재실행**: Step 1 결과를 재사용하여 비용 절약

```bash
# 테스트 워크플로우
python step1_convert.py  # 전체 변환 (무료)
python step2_analyze.py --workers 3  # 소량으로 테스트 (유료)
# 결과 확인 후
python step2_analyze.py --workers 10  # 전체 실행 (유료)
```

---

## 체크리스트

- [ ] `pip install -r requirements.txt` 실행
- [ ] `.env` 파일에 OPENAI_API_KEY 설정
- [ ] PowerPoint 파일을 `data/` 디렉토리에 배치
- [ ] Microsoft PowerPoint 설치 확인 (Windows)
- [ ] 충분한 디스크 공간 확인 (1만개 = 약 10-20GB)
- [ ] `python step1_convert.py` 실행
- [ ] `output/step1_metadata.json` 생성 확인
- [ ] `python step2_analyze.py --workers 10` 실행
- [ ] `output/processing_summary.json` 결과 확인

---

**자세한 내용은 [README.md](README.md)를 참조하세요.**
