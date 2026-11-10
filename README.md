# EchoDrive

用 Cloudflare Workers 和 OneDrive 搭建的私人网页网盘。支持新建文件夹、上传、下载、重命名和删除。

## 部署

1. 将现有的 `echodrive` Worker 连接到此仓库的 `main` 分支。
2. 根目录使用 `/`，构建命令留空，部署命令使用 `npx wrangler deploy`。
3. 在 Worker 的运行时变量中配置 `MS_CLIENT_ID`，在运行时密钥中配置 `MS_CLIENT_SECRET` 和 `MS_OWNER_ID`。
4. 在 Microsoft Entra 的 Web 重定向 URI 中配置 Worker 的 `/auth/callback` 地址。部署到其他域名时，记得同步修改源码中的 `SITE`。

网盘仅允许 `MS_OWNER_ID` 对应的微软账号访问。所有密钥通过 Cloudflare 配置。

## 首次使用

1. 打开网盘，点击 **首次绑定或重置密码**，用 `MS_OWNER_ID` 对应的微软账号授权。
2. 设置至少 12 个字符的独立网盘密码。之后打开网盘，直接输密码解锁。
3. 如果微软撤销授权，或需要重置密码，重新走首次绑定流程。

密码以加盐哈希保存，OneDrive 续期凭据在 SQLite Durable Object 中加密保存。连续输错 8 次后，需等待最早一次错误超出 10 分钟再试。
