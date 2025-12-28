#!/bin/bash
# 配置Docker使用国内镜像加速器

echo "正在配置Docker镜像加速器..."

# 创建或备份Docker daemon配置
if [ ! -d /etc/docker ]; then
    sudo mkdir -p /etc/docker
fi

# 备份现有配置
if [ -f /etc/docker/daemon.json ]; then
    sudo cp /etc/docker/daemon.json /etc/docker/daemon.json.bak
    echo "已备份现有配置到 /etc/docker/daemon.json.bak"
fi

# 配置阿里云镜像加速器
sudo tee /etc/docker/daemon.json > /dev/null <<EOF
{
  "registry-mirrors": [
    "https://registry.cn-hangzhou.aliyuncs.com",
    "https://docker.mirrors.ustc.edu.cn",
    "https://hub-mirror.c.163.com"
  ]
}
EOF

echo "Docker镜像加速器配置完成！"

# 重启Docker服务
echo "正在重启Docker服务..."
sudo systemctl daemon-reload
sudo systemctl restart docker

echo "配置完成！请重新运行 ./start.sh"
echo ""
echo "验证配置："
echo "docker info | grep -A 10 'Registry Mirrors'"

