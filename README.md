# EchoDrive

用 Cloudflare Workers 和 OneDrive 搭建的私人网页网盘。支持新建文件夹、上传、下载、重命名和删除。

## 部署

1. 将现有的 `echodrive` Worker 连接到此仓库的 `main` 分支。
2. 根目录使用 `/`，构建命令留空，部署命令使用 `npx wrangler deploy`。
3. 在 Worker 的运行时变量中配置 `MS_CLIENT_ID`，在运行时密钥中配置 `MS_CLIENT_SECRET` 和 `MS_OWNER_ID`。
4. 在 Microsoft Entra 的 Web 重定向 URI 中配置 Worker 的 `/auth/callback` 地址。部署到其他域名时，记得同步修改源码中的 `SITE`。

网盘仅允许 `MS_OWNER_ID` 对应的微软账号访问。所有密钥通过 Cloudflare 配置。
