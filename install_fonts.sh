#!/bin/bash

echo "🔧 安装中文字体支持..."

# 更新包列表
sudo apt-get update

# 安装基础字体包
sudo apt-get install -y \
    fonts-noto-cjk \
    fonts-noto-cjk-extra \
    fonts-wqy-microhei \
    fonts-wqy-zenhei \
    fonts-arphic-ukai \
    fonts-arphic-uming \
    fontconfig

# 安装 fontconfig 工具
sudo apt-get install -y fontconfig

# 刷新字体缓存
echo "♻️  刷新字体缓存..."
fc-cache -fv

# 验证中文字体
echo ""
echo "✅ 已安装的中文字体："
fc-list :lang=zh | head -n 10

echo ""
echo "✅ 字体安装完成！"
echo "请重启应用或重新运行转换任务。"
