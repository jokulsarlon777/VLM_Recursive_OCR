# VLM Recursive OCR for PowerPoint

윈도우 PC에서 OLE 방식으로 개체 삽입된 PowerPoint 파일 내의 ppt/pptx 파일을 재귀적으로 추출하고, GPT-4o Vision을 사용하여 기술 문서 내용을 분석하는 프로젝트입니다.

## 주요 기능

### 핵심 기능
- **OLE 개체 추출**: PowerPoint 파일 내에 삽입된 ppt/pptx 파일을 자동으로 추출
- **재귀적 처리**: 추출된 PowerPoint 파일에서 다시 OLE 개체를 찾아 처리
- **슬라이드 이미지 변환**: COM 자동화를 통해 각 슬라이드를 고품질 이미지로 변환
- **VLM 분석**: GPT-4o Vision API를 사용하여 슬라이드 내용 분석
- **구조화된 출력**: 문제-원인-해결방법 중심의 JSON 형식 출력

### 성능 개선 기능
- **2단계 프로세스**: 이미지 변환과 VLM 분석을 분리하여 효율성 향상
- **병렬 처리**: ThreadPoolExecutor를 사용한 VLM 분석 병렬 처리 (최대 5개 동시 처리)
- **자동 재시도**: tenacity 라이브러리를 통한 API 호출 자동 재시도 (최대 3회, exponential backoff)
- **진행 상황 표시**: tqdm을 통한 실시간 진행 상황 모니터링
- **에러 복구**: 개별 슬라이드 분석 실패 시에도 전체 프로세스 계속 진행

## 2단계 아키텍처

```
STEP 1: 이미지 변환 (Recursive)
├── PowerPoint 파일 로드
├── 모든 슬라이드를 이미지로 변환
├── OLE 개체 추출
└── 추출된 파일에 대해 재귀적으로 반복

STEP 2: VLM 분석 (Parallel)
├── 변환된 모든 이미지 수집
├── ThreadPoolExecutor로 병렬 처리 (5 workers)
├── 실패 시 자동 재시도 (exponential backoff)
└── 진행 상황 실시간 표시

STEP 3: JSON 출력 생성
├── 파일별 개별 JSON 생성
└── 전체 요약 JSON 생성
```

### 이점
- **효율성**: 이미지 변환 실패 시 VLM 분석 재실행 불필요
- **속도**: 병렬 처리로 VLM 분석 시간 대폭 단축
- **안정성**: 에러 복구 메커니즘으로 중간 실패에도 강건
- **모니터링**: 진행 상황을 실시간으로 확인 가능

## 시스템 요구사항

- Windows OS (COM 자동화 사용)
- Microsoft PowerPoint 설치 필요
- Python 3.8 이상
- 충분한 디스크 공간 (이미지 파일 저장용)

## 설치 방법

1. 저장소 클론 또는 다운로드

2. 가상환경 생성 및 활성화 (권장)
```bash
python -m venv venv
venv\Scripts\activate
```

3. 필요한 패키지 설치
```bash
pip install -r requirements.txt
```

설치되는 패키지:
- `python-pptx`: PowerPoint 파일 파싱
- `pywin32`: Windows COM 자동화
- `openai`: GPT-4o Vision API
- `python-dotenv`: 환경 변수 관리
- `Pillow`: 이미지 처리
- `tqdm`: 진행 상황 표시
- `tenacity`: 자동 재시도 로직

4. 환경 변수 설정
```bash
copy .env.example .env
```

`.env` 파일을 열어 API 키 입력:
```
OPENAI_API_KEY=your_actual_api_key_here
GPT_MODEL=gpt-4o
OUTPUT_DIR=output
TEMP_DIR=temp
```

## 사용 방법

### 기본 사용법

1. 분석할 PowerPoint 파일을 `data/` 디렉토리에 배치

2. 메인 스크립트 실행
```bash
python main.py
```

3. 결과 확인
- `output/` 디렉토리에 각 파일별 JSON 결과 생성
- `output/processing_summary.json` 전체 처리 요약 확인
- `temp/` 디렉토리에 변환된 이미지 저장 (선택적 삭제 가능)

### 실행 예시

```bash
python main.py
```

출력 예시:
```
================================================================================
VLM Recursive OCR Processor - 2-Step Architecture
================================================================================
Data directory: /path/to/data
Output directory: /path/to/output
Temp directory: /path/to/temp
================================================================================

================================================================================
STEP 1: Converting all slides to images (recursive)
Found 2 PowerPoint files
================================================================================

Processing: sample data.pptx
[Depth 0] Processing: sample data.pptx
Converting slides to images...
Converting sample data.pptx: 100%|████████████| 5/5 [00:10<00:00,  2.00s/slide]
Converted 5 slides
Checking for embedded files...
Found 1 embedded PowerPoint files
  [Depth 1] Processing: embed_file.pptx
  Converting slides to images...
  Converting embed_file.pptx: 100%|██████████| 3/3 [00:06<00:00,  2.00s/slide]
  Converted 3 slides

================================================================================
STEP 2: Analyzing all images with VLM (parallel processing)
================================================================================

Analyzing 5 slides from: sample data.pptx
Analyzing slides: 100%|████████████████████| 5/5 [00:15<00:00,  3.00s/slide]
Completed analysis for: sample data.pptx

Analyzing 3 slides from: embed_file.pptx
Analyzing slides: 100%|████████████████████| 3/3 [00:09<00:00,  3.00s/slide]
Completed analysis for: embed_file.pptx

================================================================================
STEP 3: Generating JSON output files
================================================================================

Generating JSON files: 100%|████████████████| 2/2 [00:00<00:00, 10.00file/s]

================================================================================
Processing completed successfully!
Total files processed: 2
Total images analyzed: 8
Output saved to: /path/to/output
================================================================================
```

## 프로젝트 구조

```
VLM_Recursive_OCR_260108/
├── data/                          # 입력 PowerPoint 파일 디렉토리
│   ├── embed_file.pptx
│   └── sample data.pptx
├── output/                        # JSON 출력 결과 디렉토리
│   ├── {filename}_d0_analysis.json
│   ├── {filename}_d1_analysis.json
│   └── processing_summary.json
├── temp/                          # 임시 파일 디렉토리
│   ├── {filename}_d0_slides/      # 슬라이드 이미지
│   └── {filename}_d0_embedded/    # 추출된 OLE 개체
├── .env                           # 환경 변수 (API 키)
├── .env.example                   # 환경 변수 템플릿
├── requirements.txt               # Python 패키지 목록
├── json_format.md                 # JSON 출력 형식 정의
├── config.py                      # 설정 관리
├── pptx_extractor.py             # OLE 개체 추출
├── image_converter.py            # 슬라이드 → 이미지 변환
├── vlm_analyzer.py               # GPT-4o Vision 분석 (병렬 처리)
├── main.py                        # 메인 실행 스크립트 (2-step)
└── README.md                      # 프로젝트 문서
```

## JSON 출력 형식

### 개별 파일 분석 결과
각 PowerPoint 파일마다 개별 JSON 파일이 생성됩니다:

```json
{
  "file_info": {
    "filename": "sample data.pptx",
    "parent_file": null,
    "depth": 0,
    "total_slides": 5,
    "has_embedded_files": true,
    "embedded_files": ["embed_file.pptx"],
    "processed_at": "2025-01-08T10:30:00"
  },
  "slides": [
    {
      "slide_number": 1,
      "image_filename": "sample data_slide_001.png",
      "title": "시스템 장애 분석",
      "problem_symptom": "데이터베이스 응답 시간 지연",
      "cause": "인덱스 최적화 부재",
      "countermeasure": "인덱스 재구성 및 쿼리 최적화",
      "summary": "DB 성능 저하 문제 및 해결 방안",
      "visual_references": [
        "응답 시간 추이 그래프 (최근 1주일)",
        "쿼리 실행 계획 다이어그램"
      ],
      "additional_notes": "주말 작업 권장",
      "confidence_scores": {
        "title": 0.95,
        "problem_symptom": 0.90,
        "cause": 0.85,
        "countermeasure": 0.88
      }
    }
  ]
}
```

### 전체 요약 파일
`processing_summary.json`에는 전체 처리 결과가 포함됩니다:

```json
{
  "processing_summary": {
    "total_files_processed": 2,
    "total_images_analyzed": 8,
    "processed_at": "2025-01-08T10:30:00"
  },
  "file_hierarchy": {
    "sample data_d0": {
      "filename": "sample data.pptx",
      "parent_file": null,
      "depth": 0,
      "total_slides": 5,
      "embedded_files": ["embed_file.pptx"]
    },
    "embed_file_d1": {
      "filename": "embed_file.pptx",
      "parent_file": "sample data.pptx",
      "depth": 1,
      "total_slides": 3,
      "embedded_files": []
    }
  },
  "results": {...}
}
```

## 주요 모듈 설명

### config.py
- 환경 변수 로드 및 관리
- API 키, 디렉토리 경로 설정
- GPT-4o 시스템 프롬프트 정의
- JSON 스키마 정의

### pptx_extractor.py
- `PPTXExtractor` 클래스: PowerPoint 파일에서 OLE 개체 추출
- python-pptx 라이브러리를 사용하여 파일 구조 분석
- ppt/pptx 형식의 embedded 파일을 자동으로 저장

### image_converter.py
- `SlideImageConverter` 클래스: PowerPoint 슬라이드를 이미지로 변환
- win32com을 통한 COM 자동화
- 고해상도(1920x1080) PNG 이미지 생성
- **진행 상황 표시**: tqdm 진행률 바

### vlm_analyzer.py
- `VLMAnalyzer` 클래스: GPT-4o Vision API를 통한 이미지 분석
- **병렬 처리**: ThreadPoolExecutor (최대 5 workers)
- **자동 재시도**: tenacity 데코레이터 (3회, exponential backoff)
- **진행 상황 표시**: 병렬 처리 시에도 실시간 진행률 표시
- Base64 인코딩 및 API 호출 관리
- JSON 응답 파싱 및 에러 처리

### main.py (2-Step Architecture)
- `TwoStepPPTXProcessor` 클래스: 전체 프로세스 오케스트레이션
- **Step 1**: 재귀적 파일 처리 및 이미지 변환
- **Step 2**: 병렬 VLM 분석
- **Step 3**: JSON 결과 생성 및 저장
- 파일 계층 구조 관리
- 이미지 레지스트리 관리

## 성능 최적화

### 병렬 처리
- **VLM 분석**: 최대 5개 슬라이드 동시 분석
- **처리 속도**: 순차 처리 대비 약 3-4배 빠름
- **조정 가능**: `max_vlm_workers` 파라미터로 worker 수 조정

### 에러 복구
- **자동 재시도**: API 오류 시 최대 3회 재시도
- **Exponential Backoff**: 2초 → 4초 → 8초 대기
- **부분 실패 허용**: 개별 슬라이드 실패해도 전체 프로세스 계속

### 메모리 관리
- 이미지를 temp 디렉토리에 저장 (메모리 절약)
- 처리 완료 후 선택적으로 temp 파일 삭제 가능

## 설정 커스터마이징

### VLM Workers 수 조정
```python
# main.py에서
processor = TwoStepPPTXProcessor(max_vlm_workers=10)  # 기본값: 5
```

### 이미지 해상도 조정
```python
# config.py 또는 image_converter.py에서
convert_pptx_to_images(
    pptx_path,
    output_dir,
    width=2560,  # 기본값: 1920
    height=1440  # 기본값: 1080
)
```

### 재시도 설정 조정
```python
# vlm_analyzer.py에서
@retry(
    stop=stop_after_attempt(5),  # 기본값: 3
    wait=wait_exponential(multiplier=2, min=4, max=20)  # 기본값: 1, 2, 10
)
```

## 주의사항

1. **Windows 전용**: COM 자동화는 Windows에서만 작동합니다
2. **PowerPoint 필수**: Microsoft PowerPoint가 설치되어 있어야 합니다
3. **API 비용**: OpenAI GPT-4o Vision API 사용에 따른 비용이 발생합니다
   - 병렬 처리로 인해 빠르게 많은 요청이 발생할 수 있습니다
4. **처리 시간**: 슬라이드가 많거나 embedded 파일이 많을 경우 처리 시간이 길어질 수 있습니다
5. **디스크 공간**: 모든 슬라이드가 이미지로 변환되므로 충분한 디스크 공간 필요

## 문제 해결

### PowerPoint 초기화 실패
```
RuntimeError: Failed to initialize PowerPoint
```
→ Microsoft PowerPoint가 설치되어 있는지 확인하세요

### API 키 오류
```
ValueError: OPENAI_API_KEY not found
```
→ `.env` 파일에 올바른 API 키가 설정되어 있는지 확인하세요

### API Rate Limit 오류
- 병렬 worker 수를 줄이세요: `max_vlm_workers=3`
- 자동 재시도가 활성화되어 있어 일시적인 rate limit은 자동 복구됩니다

### JSON 파싱 오류
- GPT-4o 응답이 JSON 형식이 아닐 경우 발생
- 로그를 확인하여 원본 응답 검토
- temperature 값을 더 낮춰보세요 (config.py에서 0.1 → 0.0)

### 메모리 부족
- VLM worker 수를 줄이세요
- 이미지 해상도를 낮추세요
- 처리 후 temp 파일을 자동 삭제하도록 설정하세요

## 개발자 정보

프로젝트에 대한 질문이나 개선 사항이 있으시면 이슈를 생성해 주세요.

## 라이선스

이 프로젝트는 개인 및 교육 목적으로 자유롭게 사용할 수 있습니다.
