#!/bin/bash

# 骑缝章工具 macOS应用构建脚本

echo "开始构建骑缝章工具 macOS应用..."

# 安装依赖
echo "安装依赖..."
npm install

# 构建应用
echo "构建应用..."
npm run build

echo "构建完成！应用位于 dist 目录"