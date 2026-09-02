# 📊 粤徽交付中心 · 八月数据看板

基于「粤徽交付中心八月数据看板.xlsx」构建的 Streamlit 可视化应用，展示 8 月每位员工的招聘数据。

## 快速开始（本地）

**方式一（推荐）**：双击 `start.bat`，自动在浏览器打开看板。

**方式二（命令行）**：

```bash
streamlit run app.py
```

> 数据源优先级：本机 `D:/YueHuiProject/粤徽交付中心八月数据看板.xlsx` → 仓库内同名文件（云端用）
> 依赖：`pip install -r requirements.txt`

## 云端部署（GitHub + Streamlit Cloud）

**1. 推送到 GitHub**

新建仓库后上传以下文件（或直接解压部署包到仓库根目录）：

```
app.py                        # 主程序
requirements.txt              # 依赖清单
README.md
.gitignore
.streamlit/config.toml        # 主题配置
粤徽交付中心八月数据看板.xlsx   # 数据文件（必须上传，云端依赖它）
```

```bash
git init
git add .
git commit -m "init dashboard"
git branch -M main
git remote add origin https://github.com/<你的用户名>/<仓库名>.git
git push -u origin main
```

**2. 在 Streamlit Cloud 部署**

1. 访问 https://share.streamlit.io 并用 GitHub 账号登录
2. 点击 **New app** → 选择刚推送的仓库 → Branch 选 `main` → Main file 填 `app.py`
3. 点 **Deploy**，等待 1-2 分钟构建完成
4. 部署成功后得到一个公网地址：`https://<你的应用名>.streamlit.app`，分享给任何人即可访问

> 更新数据：修改 Excel 后重新 push 到 GitHub，云端会自动重新部署
> 免费额度足够个人/小团队使用；如需私有部署可换 Hugging Face Spaces 或自有服务器

## 功能

- **侧边栏筛选**：日期范围（默认全月，可切到 8/18–8/24）、组别、人员、指标
- **📊 总览**：5 项 KPI 卡片（新增微信 / 预约 / 到场 / 合格 / 在职）、按组别汇总、每人分指标对比（每指标一张图）
- **👥 每人明细**：区间合计表（含 BOSS 账号、目标在职、达成率）+ 人 × 日期热力图
- **📈 趋势分析**：每人/全组每日趋势、多指标对比折线
- **📋 数据表**：每日明细长表 + 区间汇总 Excel 一键下载

## 数据结构说明

原表为透视布局：员工行 + 每日 5 列一组（新增微信/预约/到场/合格/在职）。应用自动解析该结构，空值按 0 处理。

## 已知说明

- 组长（王志英、魏亚飞）个人日常明细基本为 0，属数据本身情况
