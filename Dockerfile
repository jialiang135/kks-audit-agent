FROM python:3.12-slim

ENV TZ=Asia/Shanghai \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app

# 依赖极轻:仅 openpyxl(标准库 HTTP 服务,无 web 框架)
COPY pyproject.toml ./
RUN pip install "openpyxl>=3.1,<4"

COPY . .

# 审核输出目录(runs/)由运行时挂载持久化
RUN mkdir -p /app/runs
VOLUME ["/app/runs"]

EXPOSE 8080

CMD ["python", "server.py", "--host", "0.0.0.0", "--port", "8080"]
