FROM python:3.12-slim

WORKDIR /app

# 换国内 apt 源（阿里云）
RUN sed -i 's/deb.debian.org/mirrors.aliyun.com/g' /etc/apt/sources.list.d/debian.sources 2>/dev/null || \
    sed -i 's/deb.debian.org/mirrors.aliyun.com/g' /etc/apt/sources.list 2>/dev/null || true

# 系统依赖（Playwright Chromium 需要）
RUN apt-get update && apt-get install -y --no-install-recommends \
    libnss3 libnspr4 libdbus-1-3 libatk1.0-0 libatk-bridge2.0-0 \
    libcups2 libdrm2 libxkbcommon0 libatspi2.0-0 libxcomposite1 \
    libxdamage1 libxfixes3 libxrandr2 libgbm1 libpango-1.0-0 \
    libcairo2 libasound2 libx11-xcb1 fontconfig \
    && rm -rf /var/lib/apt/lists/*

# 换清华 pip 源 + 安装依赖
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt \
    -i https://pypi.tuna.tsinghua.edu.cn/simple \
    --trusted-host pypi.tuna.tsinghua.edu.cn

# Playwright 用淘宝镜像下载 Chromium
ENV PLAYWRIGHT_DOWNLOAD_HOST=https://registry.npmmirror.com/-/binary/playwright
RUN playwright install chromium

# 安装微软雅黑字体（放最后，避免缓存失效）
COPY fonts/msyh.ttc /usr/share/fonts/truetype/msyh.ttc
RUN fc-cache -fv

# 复制所有项目文件
COPY processor.py config.py ./
COPY app/ ./app/

WORKDIR /app/app

ENV PORT=45785

EXPOSE 45785

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "45785"]
