# 通用邮件清理工具 (Universal Email Cleaner)

[English](README.md) | 中文

一款功能强大且易于使用的 Windows GUI 工具，帮助管理员批量清理 Microsoft Exchange / Exchange Online 中的邮件与会议。支持 **Microsoft Graph API** 和 **EWS** 双协议，提供交互式扫描结果、多种删除模式、进度条和完善的日志系统。

---

## 功能特性

### 双协议支持

| 协议 | 适用场景 |
|------|---------|
| **Microsoft Graph API** | Exchange Online (Microsoft 365)，含全球版 & 世纪互联版 |
| **EWS (Exchange Web Services)** | Exchange Server 2010/2013/2016/2019 及 Exchange Online |

### Graph API 认证

- **自动配置 (Auto)**：自动创建 Azure AD 应用 + 自签名证书，一键连接（需全局管理员）
- **手动配置 (Manual)**：使用已有的 App ID / Tenant ID / Client Secret 或证书指纹

### EWS 认证

- **NTLM**（默认）— 适用于本地 Exchange Server
- **Basic** — 基本认证
- **OAuth2** — 现代认证（Azure AD 应用）
- **Token** — 直接使用访问令牌（支持 DPAPI 加密缓存）

### 清理目标与筛选

- **目标对象**：邮件 (Email) 或 会议 (Meeting)
- **筛选条件**：主题、发件人、正文关键字、日期范围、Message ID
- **会议范围**：单次会议、系列会议、仅已取消的会议
- **文件夹范围**：收件箱、已发送、已删除、垃圾邮件、可恢复删除项 (Recoverable Items)、全邮箱等 12 种选项
- **CSV 用户列表**：支持 `UserPrincipalName` 列的 CSV 文件
- **TXT 邮箱导入**：支持一行一个邮箱地址的纯文本文件

### 三种删除模式

| 模式 | Graph API | EWS | 说明 |
|------|-----------|-----|------|
| **删除 (可恢复)** | `DELETE` 请求 | `SoftDelete` | 邮件进入 Recoverable Items，管理员可通过 eDiscovery 恢复 |
| **移到已删除文件夹** | `POST .../move` → deleteditems | `MoveToDeletedItems` | 邮件移到 Deleted Items，用户可手动恢复 |
| **彻底删除 (不可恢复)** | `POST .../permanentDelete` | `HardDelete` / `Purge` | 永久删除，不进入 Recoverable Items |

### 交互式扫描结果

- 以「仅报告」模式扫描后，结果展示在独立选项卡
- 支持 **勾选** 单个/多个/全选/反选
- **删除模式下拉框**与任务配置页双向联动
- 删除后行变灰（成功）或变红（失败），状态显示在 Details 列
- 可加载历史 CSV 报告重新操作

### 进度与日志

- **实时进度条** + 百分比显示
- 三级日志：默认 / 高级 / 专家
- 日志与报告保存在 `%USERPROFILE%\Documents\UniversalEmailCleaner\`
- Graph Token 默认打码，专家模式下可选择记录（带安全警告）

---

## 系统要求

- Windows 10/11（推荐）
- Python 3.8+（如使用源码运行）
- 或直接下载编译好的 [.exe 文件](https://github.com/andylu1988/UniversalEmailCleaner/releases)

## 安装

### 方式一：下载 exe（推荐）

前往 [Releases](https://github.com/andylu1988/UniversalEmailCleaner/releases) 下载最新版 `.exe`，双击运行即可。

### 方式二：源码运行

```bash
git clone https://github.com/andylu1988/UniversalEmailCleaner.git
cd UniversalEmailCleaner
pip install -r requirements.txt
python universal_email_cleaner.py
```

依赖库：`requests`、`exchangelib`、`azure-identity`、`msal`（可选）、`Pillow`（可选，用于图标显示）

---

## 使用指南

### 1. 连接配置（选项卡一）

**Graph API：**
1. 选择环境：全球版 (Global) 或 世纪互联 (China)
2. 选择认证模式：Auto（一键配置）或 Manual（手动填写）
3. Auto 模式下需全局管理员权限，会自动创建 Azure AD App 和证书

**EWS：**
1. 输入服务器地址或勾选自动发现 (Autodiscover)
2. 输入管理员账号 (UPN) 和密码
3. 选择认证方式：NTLM / Basic / OAuth2 / Token
4. 选择访问类型：模拟 (Impersonation) 或 代理 (Delegate)

### 2. 任务配置（选项卡二）

1. 选择或导入邮箱列表（CSV 或 TXT 文件）
2. 设置清理目标（邮件/会议）和筛选条件
3. 选择文件夹范围
4. 选择执行选项：
   - **仅报告**：只生成 CSV 报告，不删除
   - **移到已删除文件夹**：移至 Deleted Items，用户可恢复
   - **彻底删除**：永久删除，不可恢复
5. 点击「开始扫描」或「开始清理」

### 3. 扫描结果（选项卡三）

1. 以「仅报告」模式运行后，自动跳转到此选项卡
2. 勾选要操作的项目（全选 / 取消全选 / 反选）
3. 选择删除模式（与任务配置联动）
4. 点击「删除选中项」执行删除
5. 也可点击「刷新/加载报告」导入历史 CSV

---

## 日志与报告

| 类型 | 路径 |
|------|------|
| 操作日志 | `%USERPROFILE%\Documents\UniversalEmailCleaner\app_YYYY-MM-DD.log` |
| CSV 报告 | `%USERPROFILE%\Documents\UniversalEmailCleaner\Reports\` |
| 高级日志 | `app_advanced_YYYY-MM-DD.log` |
| 专家日志 | `app_expert_YYYY-MM-DD.log` |

---

## 许可证

MIT License
