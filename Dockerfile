# OLA 弥撒 PPTX 生成工具 —— 生产镜像
# 关键：工具用系统命令 unzip / zip 解压和重打包 pptx，所以镜像里必须装上它们。
FROM node:20-slim

RUN apt-get update \
  && apt-get install -y --no-install-recommends zip unzip \
  && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# 先装依赖（利用缓存）
COPY package.json package-lock.json* ./
RUN npm install --omit=dev

# 拷贝源码 + 模板 + 前端资源
COPY . .

# 对外监听（主机会注入 PORT；HOST 必须 0.0.0.0 才能被外部访问）
ENV HOST=0.0.0.0
EXPOSE 4177

CMD ["node", "server.mjs"]
