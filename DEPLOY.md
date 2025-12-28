# 广西人身损害赔偿计算器 - Docker部署指南

本指南将帮助您将广西人身损害赔偿计算器部署到阿里云ECS服务器上。

## 前置要求

1. **阿里云ECS实例**
   - 操作系统：Ubuntu 20.04/22.04 或 CentOS 7/8
   - 至少2GB内存
   - 至少10GB磁盘空间

2. **已安装的软件**
   - Docker（版本20.10+）
   - Docker Compose（版本2.0+）

## 部署步骤

### 1. 连接到ECS服务器

使用SSH连接到您的ECS实例：

```bash
ssh root@your-server-ip
```

### 2. 安装Docker和Docker Compose

#### Ubuntu/Debian系统：

```bash
# 更新系统包
apt-get update

# 安装Docker
curl -fsSL https://get.docker.com -o get-docker.sh
sh get-docker.sh

# 启动Docker服务
systemctl start docker
systemctl enable docker

# 安装Docker Compose
curl -L "https://github.com/docker/compose/releases/latest/download/docker-compose-$(uname -s)-$(uname -m)" -o /usr/local/bin/docker-compose
chmod +x /usr/local/bin/docker-compose
```

#### CentOS系统：

```bash
# 更新系统包
yum update -y

# 安装Docker
yum install -y docker
systemctl start docker
systemctl enable docker

# 安装Docker Compose
curl -L "https://github.com/docker/compose/releases/latest/download/docker-compose-$(uname -s)-$(uname -m)" -o /usr/local/bin/docker-compose
chmod +x /usr/local/bin/docker-compose
```

### 3. 配置Docker镜像加速器（重要！）

**这是解决网络连接问题的关键步骤！**

```bash
# 运行配置脚本
sudo ./setup-docker-mirror.sh

# 或者手动配置
sudo mkdir -p /etc/docker
sudo tee /etc/docker/daemon.json > /dev/null <<EOF
{
  "registry-mirrors": [
    "https://registry.cn-hangzhou.aliyuncs.com",
    "https://docker.mirrors.ustc.edu.cn",
    "https://hub-mirror.c.163.com"
  ]
}
EOF

sudo systemctl daemon-reload
sudo systemctl restart docker

# 验证配置
docker info | grep -A 10 'Registry Mirrors'
```

### 4. 上传项目文件到服务器

#### 方法1：使用Git（推荐）

```bash
# 在服务器上克隆项目
cd /opt
git clone your-repository-url guangxi_compensation_calculator
cd guangxi_compensation_calculator
```

#### 方法2：使用SCP上传

在本地机器上执行：

```bash
scp -r /path/to/project root@your-server-ip:/opt/guangxi_compensation_calculator/
```

然后在服务器上：

```bash
cd /opt/guangxi_compensation_calculator
```

### 5. 构建和启动Docker容器

```bash
# 创建临时目录
mkdir -p temp

# 构建Docker镜像
docker-compose build

# 启动服务
docker-compose up -d

# 查看日志
docker-compose logs -f
```

### 6. 配置防火墙

确保ECS安全组和系统防火墙允许5000端口：

```bash
# Ubuntu/Debian
ufw allow 5000/tcp
ufw reload

# CentOS
firewall-cmd --permanent --add-port=5000/tcp
firewall-cmd --reload
```

### 7. 配置Nginx反向代理（可选但推荐）

安装Nginx：

```bash
# Ubuntu/Debian
apt-get install -y nginx

# CentOS
yum install -y nginx
```

创建Nginx配置文件：

```bash
vim /etc/nginx/sites-available/compensation-calculator
```

添加以下内容：

```nginx
server {
    listen 80;
    server_name your-domain.com;  # 替换为您的域名或IP

    location / {
        proxy_pass http://127.0.0.1:5000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

启用配置：

```bash
# Ubuntu/Debian
ln -s /etc/nginx/sites-available/compensation-calculator /etc/nginx/sites-enabled/
nginx -t
systemctl restart nginx

# CentOS
cp /etc/nginx/sites-available/compensation-calculator /etc/nginx/conf.d/
nginx -t
systemctl restart nginx
```

### 8. 配置SSL证书（可选，推荐用于生产环境）

使用Let's Encrypt免费SSL证书：

```bash
# 安装Certbot
apt-get install -y certbot python3-certbot-nginx  # Ubuntu/Debian
# 或
yum install -y certbot python3-certbot-nginx  # CentOS

# 获取证书
certbot --nginx -d your-domain.com

# 自动续期
certbot renew --dry-run
```

## 访问应用

部署完成后，您可以通过以下方式访问：

- **直接访问**：`http://your-server-ip:5000`
- **通过Nginx**：`http://your-domain.com` 或 `https://your-domain.com`（如果配置了SSL）

## 常用管理命令

### 查看容器状态

```bash
docker-compose ps
```

### 查看日志

```bash
# 实时日志
docker-compose logs -f

# 最近100行日志
docker-compose logs --tail=100
```

### 重启服务

```bash
docker-compose restart
```

### 停止服务

```bash
docker-compose stop
```

### 启动服务

```bash
docker-compose start
```

### 更新应用

```bash
# 拉取最新代码
git pull  # 如果使用Git

# 重新构建并启动
docker-compose up -d --build
```

### 查看容器资源使用情况

```bash
docker stats guangxi_compensation_calculator
```

## 故障排查

### 1. Docker镜像拉取失败（网络超时）

**问题**：`failed to resolve source metadata for docker.io/library/python:3.11-slim`

**解决方案**：
```bash
# 1. 配置Docker镜像加速器（必须！）
sudo ./setup-docker-mirror.sh

# 2. 验证配置
docker info | grep -A 10 'Registry Mirrors'

# 3. 重新构建
docker-compose build
```

### 2. 容器无法启动

检查日志：

```bash
docker-compose logs
```

### 3. 端口被占用

检查端口占用：

```bash
netstat -tulpn | grep 5000
```

修改`docker-compose.yml`中的端口映射。

### 4. 无法访问应用

- 检查ECS安全组规则是否开放5000端口
- 检查系统防火墙设置
- 检查容器是否正常运行：`docker-compose ps`

### 5. Word文档导出失败

检查临时目录权限：

```bash
mkdir -p temp
chmod 777 temp
```

### 6. pip安装依赖失败

如果pip安装Python包时失败，Dockerfile已配置使用阿里云镜像源。如果仍有问题，可以手动测试：

```bash
docker run -it python:3.11-slim bash
pip config set global.index-url https://mirrors.aliyun.com/pypi/simple/
pip install flask python-docx
```

## 性能优化建议

1. **使用Nginx反向代理**：提高性能和安全性
2. **配置Gunicorn**：生产环境建议使用Gunicorn替代Flask内置服务器
3. **设置资源限制**：在`docker-compose.yml`中添加资源限制
4. **定期清理**：定期清理临时文件和日志

## 安全建议

1. **修改SECRET_KEY**：在`app.py`中修改`SECRET_KEY`
2. **使用HTTPS**：配置SSL证书
3. **限制访问**：使用防火墙限制访问IP
4. **定期更新**：保持Docker镜像和依赖库更新

## 备份和恢复

### 备份

```bash
# 备份应用代码
tar -czf backup-$(date +%Y%m%d).tar.gz /opt/guangxi_compensation_calculator
```

### 恢复

```bash
# 解压备份
tar -xzf backup-YYYYMMDD.tar.gz -C /
# 重新构建和启动
cd /opt/guangxi_compensation_calculator
docker-compose up -d --build
```

## 技术支持

如有问题，请联系：
- 广西瀛桂律师事务所 唐学智律师
- 联系电话：18078374299

## 更新日志

### v1.0.0 (2025)
- 初始Docker部署版本
- 支持Web界面访问
- 支持Word文档导出
- 配置国内镜像加速器支持
