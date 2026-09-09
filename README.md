# 手术室打卡统计（daka / daka2）

手术室微信打卡 Excel 统计工具，处理打卡记录、按区域/班次分组、按上班时长标色，生成统计表格图片。包含两个独立的 Docker 网页版：

| | **daka**（一区） | **daka2**（二区） |
|---|---|---|
| 目录 | `app/` + 根目录 `Dockerfile` | `app2/`（自包含） |
| 端口 | 45785 | 45786 |
| 访问前缀 | `/daka` | `/daka2` |
| 数据格式 | 手填下班时间（模糊解析） | 上/下班完整日期时间直接相减 |
| 分组 | 按楼栋/区域 | 白班 → 接班 → 值班（按明日安排） |
| 详见 | `云端部署方案.md` | `app2/部署说明.md` |

两个版本可部署在同一台服务器，端口、容器名、登录 Cookie 相互独立。

## 目录结构

```
├── Dockerfile / docker-compose.yml    # daka 镜像与编排
├── processor.py / config.py           # daka 数据处理与配置
├── requirements.txt
├── app/                               # daka 后端 + 前端页面
│   ├── main.py
│   └── static/（index.html / login.html）
├── app2/                              # daka2 完整独立应用
│   ├── main.py / processor2.py / config2.py
│   ├── Dockerfile / docker-compose.yml / requirements.txt
│   ├── static/
│   └── 部署说明.md
└── 云端部署方案.md                     # daka 宝塔部署文档
```

## 快速开始

### 1. 准备字体（构建镜像必需，因版权不入库）

把微软雅黑 `msyh.ttc` 放到项目根目录 `fonts/` 下：

```
fonts/msyh.ttc
```

### 2. 配置密码

```bash
# daka
cp .env.example .env        # 编辑 SITE_PASSWORD

# daka2
cd app2 && cp .env.example .env && cd ..
```

密码只通过 `SITE_PASSWORD` 环境变量注入，代码中不保存；未设置时应用可启动但无法登录。

### 3. 构建启动

```bash
# daka（端口 45785）
docker compose up -d --build

# daka2（端口 45786，在项目根目录执行）
docker compose -f app2/docker-compose.yml up -d --build
```

### 4. 访问

- daka：`http://IP:45785/daka/`
- daka2：`http://IP:45786/daka2/`

同一域名下可用 Nginx 按 `/daka`、`/daka2` 前缀反代分流（配置示例见 `app2/部署说明.md`）。

## 使用流程

1. 打开页面，输入密码登录
2. 选择微信导出的打卡 Excel（文件名 `YYYY-MM-DD至YYYY-MM-DD_打卡记录.xlsx`）上传
3. 「开始处理」→ 预览并下载 `YYYYMMDD下班统计.png`
4. 数据有误时「修改表格」：双击编辑、新增行，保存后重新处理

上传与产物文件 7 天自动清理（容器启动时执行）。

## 注意

- 打卡数据（`.xlsx`）与统计图（`.png`）含个人信息，已通过 `.gitignore` 排除，**不要提交到仓库**
- 服务器内存建议 ≥ 2GB/实例（Playwright Chromium 截图）
