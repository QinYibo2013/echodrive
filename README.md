# EchoDrive

用 Cloudflare Workers 和 OneDrive 搭建的私人网页网盘。支持新建文件夹、上传、下载、重命名和删除。

## 部署

1. 将现有的 `echodrive` Worker 连接到此仓库的 `main` 分支。
2. 根目录使用 `/`，构建命令留空，部署命令使用 `npx wrangler deploy`。
3. 在 Worker 的运行时变量中配置 `MS_CLIENT_ID`，在运行时密钥中配置 `MS_CLIENT_SECRET` 、`MS_OWNER_ID` 和 `ECHO_PASSWORD`。
4. 在 Microsoft Entra 的 Web 重定向 URI 中配置 Worker 的 `/auth/callback` 地址。部署到其他域名时，记得同步修改源码中的 `SITE`。

网盘仅允许 `MS_OWNER_ID` 对应的微软账号访问。所有密钥通过 Cloudflare 配置。

## 首次使用

1. 在 Cloudflare 的运行时密钥中设置独立的 `ECHO_PASSWORD`，建议至少 12 个字符。
2. 如果 OneDrive 尚未绑定，打开网盘，点击 **首次绑定或重新授权 OneDrive**，用网盘主人的微软账号授权。
3. 之后打开网盘直接输入 `ECHO_PASSWORD`。改密码时更新 Cloudflare 密钥并部署，旧会话随后失效。

OneDrive 续期凭据在 SQLite Durable Object 中加密保存。十分钟内输错八次后，需等待最早一次错误超出十分钟再试。
