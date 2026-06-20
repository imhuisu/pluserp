import multiprocessing

# SQLite 동시쓰기 충돌 방지: worker는 1~2개로 제한
# PostgreSQL로 전환 후: workers = multiprocessing.cpu_count() * 2 + 1 로 변경
workers = 2
worker_class = "gthread"   # 스레드 기반 → 동시 요청 처리
threads = 4                # worker당 스레드 수 (총 동시처리: 2 * 4 = 8)

bind = "0.0.0.0:5000"
timeout = 120              # 파일 업로드/PDF 생성 등 긴 작업 고려
keepalive = 5
max_requests = 1000        # 메모리 누수 방지: 1000 요청마다 worker 재시작
max_requests_jitter = 100  # 동시 재시작 방지용 랜덤 편차

accesslog = "-"            # stdout으로 접근 로그
errorlog = "-"
loglevel = "info"
