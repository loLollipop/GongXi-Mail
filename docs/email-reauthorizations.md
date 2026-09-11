# Microsoft 邮箱批量重新授权

超级管理员在「邮箱管理」后的「批量重新授权」页面启动队列。每个账号仍需在 Microsoft 页面确认一次；应用不会代替用户输入密码，也不会自动打开下一账号的标签页。验证码只在当前浏览器内存和受权限保护的会话接口中展示，device code 加密存储，access token 仅在服务端请求内存中使用，refresh token 核验成功后加密保存。

## 部署与数据库

此仓库的 Docker 启动命令使用 `prisma db push --skip-generate`，没有迁移历史。本次保持这个部署约定：在更新服务之前备份数据库，生成新 Prisma Client，并在维护窗口执行 `prisma db push`。不得使用 `--accept-data-loss`。本次 schema 变更仅增加 `email_accounts.token_version`（默认 0）、授权会话表、枚举、外键和索引，不删除现有数据。不要直接在已有、未建立迁移基线的数据库执行 `prisma migrate deploy`。

现有邮箱从版本 0 开始；所有新版本的手工修改、导入和 token 轮换都会增加版本。部署时停止旧实例及其定时刷新任务，完成 schema 更新后只运行新实例；旧二进制不遵守版本 CAS，会破坏并发保护。回退代码时先停止授权和 token 刷新任务，确认没有在途请求；新增列/表可保留。

数据库唯一的可空 `active_email_id` 保证同一邮箱只有一个活动会话，不依赖单进程锁。轮询通过数据库状态与 claim 比较更新串行化；token 写入和会话成功状态位于同一个事务，并对邮箱版本、身份和禁用状态进行校验。请求超时 15 秒，轮询占用超过 60 秒后会话失败，需重新开始，避免重复兑换已消费的设备码。

到期会话在访问授权接口时清理，终态立即清除设备码与用户验证码；终态记录保留 30 天并按需删除。无后台轮询 worker。页面刷新可恢复仍活动的会话，点击继续后恢复轮询。本轮成功/跳过/失败计数在刷新队列时重置。

## Microsoft 应用配置

使用各邮箱当前 `clientId`，应用必须支持个人 Microsoft 账号与 public client/device-code flow。默认请求 `openid profile offline_access https://graph.microsoft.com/User.Read https://graph.microsoft.com/Mail.ReadWrite`，仅包含 Graph 资源且不同时请求 `Mail.Read`。token 响应必须明确包含大小写不敏感的 `Mail.ReadWrite` 或完整 Graph URI scope，否则以 `GRAPH_MAIL_READWRITE_SCOPE_MISSING` 失败且不保存 refresh token。`GET /me` 的 `mail` 或 `userPrincipalName` 必须与目标邮箱大小写不敏感相等；无法证明别名归属时拒绝保存。客户端配置或租户策略不允许 device flow 时页面显示失败，不能以跳过身份核验的方式绕过。

本功能恢复 Graph 读写授权，因为现有 `processMailbox` 会删除邮件，必须具备 `Mail.ReadWrite`。本功能不会申请 IMAP 权限；`IMAP_ONLY` 分组会在候选列表明确标为不支持，服务端也会以 `REAUTHORIZATION_STRATEGY_UNSUPPORTED` 拒绝启动。此类账号需先改为可回退到 Graph 的策略，或另行完成 IMAP 授权。自动 token 刷新会跳过明确要求用户重新登录的账号，单账号手工刷新仍可重试。

## 上线核验

本地自动测试使用模拟 Microsoft 响应和数据库操作，不连接真实账号。上线前在独立测试库检验 `prisma db push`、同邮箱并发启动唯一性、跨实例轮询 CAS、取消与兑换竞争，以及 Microsoft 应用的 public client 配置；再使用获授权的测试邮箱验证真实登录、错误账号拒绝、成功后的 Graph 读信和页面刷新恢复。不要把令牌或 device code 放进测试日志。
