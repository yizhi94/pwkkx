# 将 pwkkx 推送到 GitHub (yizhi94)

本地已完成：首次提交、远程 `origin` 已指向 `https://github.com/yizhi94/pwkkx.git`。

你只需在本地完成一次**推送**（需本机已登录 GitHub）。

## 1. 在 GitHub 上新建仓库（若尚未创建）

1. 打开 https://github.com/new  
2. **Repository name** 填：`pwkkx`  
3. **Owner** 选：`yizhi94`  
4. 选择 Public 或 Private，**不要**勾选 “Add a README”  
5. 点击 **Create repository**

## 2. 在本地推送

在项目根目录执行：

```bash
cd /mnt/d/pwkkx
git push -u origin main
```

- 若使用 **HTTPS** 且未配置凭据，会提示输入用户名与密码；密码需使用 [Personal Access Token](https://github.com/settings/tokens)，不要用登录密码。  
- 若已配置 **SSH 公钥**，可改为用 SSH 地址再推送：

```bash
git remote set-url origin git@github.com:yizhi94/pwkkx.git
git push -u origin main
```

成功后在浏览器打开：**https://github.com/yizhi94/pwkkx** 即可看到仓库内容。
