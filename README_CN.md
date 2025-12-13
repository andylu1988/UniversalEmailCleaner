# 通用邮件清理工具 (Universal Email Cleaner)

[English README](README.md)

一款功能强大且易于使用的 GUI 工具，旨在帮助管理员和用户清理 Microsoft Exchange Server 和 Exchange Online 中的邮件与会议。

## 主要功能

*   **双协议支持**：
    *   **EWS (Exchange Web Services)**：支持 Exchange Server 2010, 2013, 2016, 2019 以及 Exchange Online。
    *   **Microsoft Graph API**：Exchange Online (Microsoft 365) 的现代标准接口。
*   **灵活的 Graph API 配置**：
    *   **多环境支持**：完美支持 **全球版 (Global)** 和 **世纪互联 (China 21Vianet)** 云环境。
    *   **认证模式**：
        *   **自动配置 (Auto)**：自动创建 Azure AD 应用程序和自签名证书，实现一键连接。
        *   **手动配置 (Manual)**：支持使用现有的 App ID、Tenant ID 和 Client Secret 进行连接。
*   **全面的清理选项**：
    *   **目标对象**：清理邮件或会议。
    *   **筛选条件**：支持按主题、发件人、正文关键字、日期范围和 Message ID 进行筛选。
    *   **会议范围**：支持处理单次会议、系列会议或仅清理已取消的会议。
*   **安全机制**：
    *   **仅报告模式 (Report Only)**：生成包含待删除项目的 CSV 报告，而不执行实际删除操作。
    *   **日志记录**：提供详细的操作日志用于审计和排错。

## 系统要求

*   Windows 操作系统 (推荐)
*   Python 3.8+
*   依赖库：`requests`, `msal`, `azure-identity`, `exchangelib`, `tk`

## 安装说明

1.  克隆仓库或下载源代码。
2.  安装所需的依赖库：
    ```bash
    pip install -r requirements.txt
    ```

## 使用指南

1.  运行应用程序：
    ```bash
    python universal_email_cleaner.py
    ```
    *或者直接运行编译好的 `.exe` 文件。*

2.  **连接配置**：
    *   选择 **EWS** 或 **Graph API**。
    *   **Graph API 设置**：
        *   选择环境：**全球版 (Global)** 或 **世纪互联 (China)**。
        *   选择配置方式：**自动 (Auto)** (需要全局管理员权限以进行初始化) 或 **手动 (Manual)** (使用 Client Secret)。
    *   **EWS 设置**：
        *   输入服务器地址 (或使用自动发现)、管理员账号 (UPN) 和密码。
        *   选择访问类型：**模拟 (Impersonation)** (推荐管理员使用) 或 **代理 (Delegate)**。

3.  **任务配置**：
    *   选择包含用户邮箱列表的 CSV 文件 (需包含 `UserPrincipalName` 列)。
    *   设置清理筛选条件 (主题、发件人、日期等)。
    *   建议先勾选 **仅生成报告 (Report Only)** 进行测试。

4.  **执行**：点击“开始清理”按钮开始处理。

## 许可证

MIT License
