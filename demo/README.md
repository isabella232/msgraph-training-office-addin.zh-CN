---
ms.openlocfilehash: fed56663591ac36e4defcae3a8d0a222f86e3026
ms.sourcegitcommit: 24bb4b3df6a035806a58b609e1d8078ac5505fa1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/17/2021
ms.locfileid: "50274231"
---
# <a name="how-to-run-the-completed-project"></a><span data-ttu-id="4b8d2-101">如何运行已完成的项目</span><span class="sxs-lookup"><span data-stu-id="4b8d2-101">How to run the completed project</span></span>

## <a name="prerequisites"></a><span data-ttu-id="4b8d2-102">先决条件</span><span class="sxs-lookup"><span data-stu-id="4b8d2-102">Prerequisites</span></span>

<span data-ttu-id="4b8d2-103">若要在此文件夹中运行已完成的项目，需要执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="4b8d2-103">To run the completed project in this folder, you need the following:</span></span>

- <span data-ttu-id="4b8d2-104">[Node.js](https://nodejs.org) 计算机上安装的Node.js和 [一](https://yarnpkg.com/) 线。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-104">[Node.js](https://nodejs.org) and [Yarn](https://yarnpkg.com/) installed on your development machine.</span></span> <span data-ttu-id="4b8d2-105"> (：本教程是使用 Node 版本 14.15.0 和一线 1.22.0 编写的。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-105">(**Note:** This tutorial was written with Node version 14.15.0 and Yarn version 1.22.0.</span></span> <span data-ttu-id="4b8d2-106">本指南中的步骤可能与其他版本一起运行，但尚未测试。) </span><span class="sxs-lookup"><span data-stu-id="4b8d2-106">The steps in this guide may work with other versions, but that has not been tested.)</span></span>
- <span data-ttu-id="4b8d2-107">具有邮箱的个人 Microsoft 帐户Outlook.com Microsoft 工作或学校帐户。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-107">Either a personal Microsoft account with a mailbox on Outlook.com, or a Microsoft work or school account.</span></span>

<span data-ttu-id="4b8d2-108">如果你没有 Microsoft 帐户，则有两种获取免费帐户的选项：</span><span class="sxs-lookup"><span data-stu-id="4b8d2-108">If you don't have a Microsoft account, there are a couple of options to get a free account:</span></span>

- <span data-ttu-id="4b8d2-109">你可以 [注册新的个人 Microsoft 帐户](https://signup.live.com/signup?wa=wsignin1.0&rpsnv=12&ct=1454618383&rver=6.4.6456.0&wp=MBI_SSL_SHARED&wreply=https://mail.live.com/default.aspx&id=64855&cbcxt=mai&bk=1454618383&uiflavor=web&uaid=b213a65b4fdc484382b6622b3ecaa547&mkt=E-US&lc=1033&lic=1)。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-109">You can [sign up for a new personal Microsoft account](https://signup.live.com/signup?wa=wsignin1.0&rpsnv=12&ct=1454618383&rver=6.4.6456.0&wp=MBI_SSL_SHARED&wreply=https://mail.live.com/default.aspx&id=64855&cbcxt=mai&bk=1454618383&uiflavor=web&uaid=b213a65b4fdc484382b6622b3ecaa547&mkt=E-US&lc=1033&lic=1).</span></span>
- <span data-ttu-id="4b8d2-110">你可以 [注册 Office 365 开发人员计划](https://developer.microsoft.com/office/dev-program) ，获取免费的 Office 365 订阅。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-110">You can [sign up for the Office 365 Developer Program](https://developer.microsoft.com/office/dev-program) to get a free Office 365 subscription.</span></span>

## <a name="register-a-web-application-with-the-azure-active-directory-admin-center"></a><span data-ttu-id="4b8d2-111">向 Azure Active Directory 管理中心注册 Web 应用程序</span><span class="sxs-lookup"><span data-stu-id="4b8d2-111">Register a web application with the Azure Active Directory admin center</span></span>

1. <span data-ttu-id="4b8d2-112">打开浏览器，并转到 [Azure Active Directory 管理中心](https://aad.portal.azure.com)。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-112">Open a browser and navigate to the [Azure Active Directory admin center](https://aad.portal.azure.com).</span></span> <span data-ttu-id="4b8d2-113">使用 **个人帐户**（亦称为“Microsoft 帐户”）或 **工作或学校帐户** 登录。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-113">Login using a **personal account** (aka: Microsoft Account) or **Work or School Account**.</span></span>

1. <span data-ttu-id="4b8d2-114">选择左侧导航栏中的“**Azure Active Directory**”，再选择“**管理**”下的“**应用注册**”。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-114">Select **Azure Active Directory** in the left-hand navigation, then select **App registrations** under **Manage**.</span></span>

    ![<span data-ttu-id="4b8d2-115">应用注册的屏幕截图</span><span class="sxs-lookup"><span data-stu-id="4b8d2-115">A screenshot of the App registrations</span></span> ](/tutorial/images/app-registrations.png)

1. <span data-ttu-id="4b8d2-116">选择“新注册”。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-116">Select **New registration**.</span></span> <span data-ttu-id="4b8d2-117">在“注册应用”页上，按如下方式设置值。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-117">On the **Register an application** page, set the values as follows.</span></span>

    - <span data-ttu-id="4b8d2-118">将“名称”设置为“`Office Add-in Graph Tutorial`”。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-118">Set **Name** to `Office Add-in Graph Tutorial`.</span></span>
    - <span data-ttu-id="4b8d2-119">将“受支持的帐户类型”设置为“任何组织目录中的帐户和个人 Microsoft 帐户”。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-119">Set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts**.</span></span>
    - <span data-ttu-id="4b8d2-120">在“重定向 URI”下，将第一个下拉列表设置为“`Single-page application (SPA)`”，并将值设置为“`https://localhost:3000/consent.html`”。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-120">Under **Redirect URI**, set the first drop-down to `Single-page application (SPA)` and set the value to `https://localhost:3000/consent.html`.</span></span>

    !["注册应用程序"页的屏幕截图](/tutorial/images/register-an-app.png)

1. <span data-ttu-id="4b8d2-122">选择“**注册**”。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-122">Select **Register**.</span></span> <span data-ttu-id="4b8d2-123">在 **"Office 外接程序图形** 教程"页上，复制应用程序应用程序 (客户端) **ID** 并保存它，下一步将需要它。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-123">On the **Office Add-in Graph Tutorial** page, copy the value of the **Application (client) ID** and save it, you will need it in the next step.</span></span>

    ![新应用注册的应用程序 ID 屏幕截图](/tutorial/images/application-id.png)

1. <span data-ttu-id="4b8d2-125">选择“管理”下的“证书和密码”。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-125">Select **Certificates & secrets** under **Manage**.</span></span> <span data-ttu-id="4b8d2-126">选择“新客户端密码”按钮。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-126">Select the **New client secret** button.</span></span> <span data-ttu-id="4b8d2-127">在 Description 中 **输入** 值，然后选择"过期"选项之 **一，然后选择**"**添加"。**</span><span class="sxs-lookup"><span data-stu-id="4b8d2-127">Enter a value in **Description** and select one of the options for **Expires** and select **Add**.</span></span>

1. <span data-ttu-id="4b8d2-128">离开此页前，先复制客户端密码值。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-128">Copy the client secret value before you leave this page.</span></span> <span data-ttu-id="4b8d2-129">将在下一步中用到它。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-129">You will need it in the next step.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="4b8d2-130">此客户端密码不会再次显示，所以请务必现在就复制它。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-130">This client secret is never shown again, so make sure you copy it now.</span></span>

1. <span data-ttu-id="4b8d2-131">在 **"管理"下\*\*\*\*选择** API 权限，然后选择 **"添加权限"。**</span><span class="sxs-lookup"><span data-stu-id="4b8d2-131">Select **API permissions** under **Manage**, then select **Add a permission**.</span></span>

1. <span data-ttu-id="4b8d2-132">选择 **Microsoft Graph，** 然后选择 **委派权限**。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-132">Select **Microsoft Graph**, then **Delegated permissions**.</span></span>

1. <span data-ttu-id="4b8d2-133">选择以下权限，然后选择 **"添加权限"。**</span><span class="sxs-lookup"><span data-stu-id="4b8d2-133">Select the following permissions, then select **Add permissions**.</span></span>

    - <span data-ttu-id="4b8d2-134">**Calendars.ReadWrite** - 这将允许应用读取和写入用户的日历。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-134">**Calendars.ReadWrite** - this will allow the app to read and write to the user's calendar.</span></span>
    - <span data-ttu-id="4b8d2-135">**MailboxSettings.Read** - 这将允许应用从用户的邮箱设置获取用户的时区。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-135">**MailboxSettings.Read** - this will allow the app to get the user's time zone from their mailbox settings.</span></span>

    ![已配置权限的屏幕截图](/tutorial/images/configured-permissions.png)

## <a name="configure-office-add-in-single-sign-on"></a><span data-ttu-id="4b8d2-137">配置 Office 外接程序单一登录</span><span class="sxs-lookup"><span data-stu-id="4b8d2-137">Configure Office Add-in single sign-on</span></span>

<span data-ttu-id="4b8d2-138">更新应用注册以支持 Office 外接程序[单一登录 (SSO) 。 ](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins)</span><span class="sxs-lookup"><span data-stu-id="4b8d2-138">Update the app registration to support [Office Add-in single sign-on (SSO)](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins).</span></span>

1. <span data-ttu-id="4b8d2-139">选择 **"公开 API"。**</span><span class="sxs-lookup"><span data-stu-id="4b8d2-139">Select **Expose an API**.</span></span> <span data-ttu-id="4b8d2-140">在此 **API 定义的** 范围部分中，选择 **"添加范围"。**</span><span class="sxs-lookup"><span data-stu-id="4b8d2-140">In the **Scopes defined by this API** section, select **Add a scope**.</span></span> <span data-ttu-id="4b8d2-141">当系统提示设置应用程序 **ID URI** 时，将值设置为 `api://localhost:3000/YOUR_APP_ID_HERE` ， `YOUR_APP_ID_HERE` 以应用程序 ID 替换。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-141">When prompted to set an **Application ID URI**, set the value to `api://localhost:3000/YOUR_APP_ID_HERE`, replacing `YOUR_APP_ID_HERE` with the application ID.</span></span> <span data-ttu-id="4b8d2-142">选择 **"保存"并继续**。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-142">Choose **Save and continue**.</span></span>

1. <span data-ttu-id="4b8d2-143">按如下所示填写字段，然后选择"**添加范围"。**</span><span class="sxs-lookup"><span data-stu-id="4b8d2-143">Fill in the fields as follows and select **Add scope**.</span></span>

    - <span data-ttu-id="4b8d2-144">**范围名称：**`access_as_user`</span><span class="sxs-lookup"><span data-stu-id="4b8d2-144">**Scope name:** `access_as_user`</span></span>
    - <span data-ttu-id="4b8d2-145">**谁可以同意？：管理员和用户**</span><span class="sxs-lookup"><span data-stu-id="4b8d2-145">**Who can consent?: Admins and users**</span></span>
    - <span data-ttu-id="4b8d2-146">**管理员同意显示名称：**`Access the app as the user`</span><span class="sxs-lookup"><span data-stu-id="4b8d2-146">**Admin consent display name:** `Access the app as the user`</span></span>
    - <span data-ttu-id="4b8d2-147">**管理员同意说明：**`Allows Office Add-ins to call the app's web APIs as the current user.`</span><span class="sxs-lookup"><span data-stu-id="4b8d2-147">**Admin consent description:** `Allows Office Add-ins to call the app's web APIs as the current user.`</span></span>
    - <span data-ttu-id="4b8d2-148">**用户同意显示名称：**`Access the app as you`</span><span class="sxs-lookup"><span data-stu-id="4b8d2-148">**User consent display name:** `Access the app as you`</span></span>
    - <span data-ttu-id="4b8d2-149">**用户同意说明：**`Allows Office Add-ins to call the app's web APIs as you.`</span><span class="sxs-lookup"><span data-stu-id="4b8d2-149">**User consent description:** `Allows Office Add-ins to call the app's web APIs as you.`</span></span>
    - <span data-ttu-id="4b8d2-150">**状态：已启用**</span><span class="sxs-lookup"><span data-stu-id="4b8d2-150">**State: Enabled**</span></span>

    !["添加范围"表单的屏幕截图](/tutorial/images/add-scope.png)

1. <span data-ttu-id="4b8d2-152">在"**授权客户端应用程序**"部分，选择 **"添加客户端应用程序"。**</span><span class="sxs-lookup"><span data-stu-id="4b8d2-152">In the **Authorized client applications** section, select **Add a client application**.</span></span> <span data-ttu-id="4b8d2-153">从以下列表中输入客户端 ID，在"授权范围"下启用范围，然后选择"**添加应用程序"。**</span><span class="sxs-lookup"><span data-stu-id="4b8d2-153">Enter a client ID from the following list, enable the scope under **Authorized scopes**, and select **Add application**.</span></span> <span data-ttu-id="4b8d2-154">对列表中的每个客户端 ID 重复此过程。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-154">Repeat this process for each of the client IDs in the list.</span></span>

    - <span data-ttu-id="4b8d2-155">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span><span class="sxs-lookup"><span data-stu-id="4b8d2-155">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span></span>
    - <span data-ttu-id="4b8d2-156">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)</span><span class="sxs-lookup"><span data-stu-id="4b8d2-156">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)</span></span>
    - <span data-ttu-id="4b8d2-157">`57fb890c-0dab-4253-a5e0-7188c88b2bb4`（Office 网页版）</span><span class="sxs-lookup"><span data-stu-id="4b8d2-157">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office on the web)</span></span>
    - <span data-ttu-id="4b8d2-158">`08e18876-6177-487e-b8b5-cf950c1e598c`（Office 网页版）</span><span class="sxs-lookup"><span data-stu-id="4b8d2-158">`08e18876-6177-487e-b8b5-cf950c1e598c` (Office on the web)</span></span>

## <a name="install-development-certificates"></a><span data-ttu-id="4b8d2-159">安装开发证书</span><span class="sxs-lookup"><span data-stu-id="4b8d2-159">Install development certificates</span></span>

1. <span data-ttu-id="4b8d2-160">运行以下命令，为加载项生成和安装开发证书。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-160">Run the following command to generate and install development certificates for your add-in.</span></span>

    ```Shell
    npx office-addin-dev-certs install
    ```

    <span data-ttu-id="4b8d2-161">如果系统提示确认，请确认操作。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-161">If prompted for confirmation, confirm the actions.</span></span> <span data-ttu-id="4b8d2-162">命令完成后，你将看到如下所示的输出。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-162">Once the command completes, you will see output similar to the following.</span></span>

    ```Shell
    You now have trusted access to https://localhost.
    Certificate: <path>\localhost.crt
    Key: <path>\localhost.key
    ```

1. <span data-ttu-id="4b8d2-163">将路径复制到 localhost.crt 和 localhost.key，下一步将需要它们。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-163">Copy the paths to localhost.crt and localhost.key, you'll need them in the next step.</span></span>

## <a name="update-the-manifest"></a><span data-ttu-id="4b8d2-164">更新清单</span><span class="sxs-lookup"><span data-stu-id="4b8d2-164">Update the manifest</span></span>

1. <span data-ttu-id="4b8d2-165">打开 **manifest.xml** 文件，然后进行以下更改。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-165">Open the **manifest.xml** file and make the following changes.</span></span>
    1. <span data-ttu-id="4b8d2-166">替换为 `NEW_GUID_HERE` 新的 GUID，如 `b4fa03b8-1eb6-4e8b-a380-e0476be9e019` 。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-166">Replace `NEW_GUID_HERE` with a new GUID, like `b4fa03b8-1eb6-4e8b-a380-e0476be9e019`.</span></span>
    1. <span data-ttu-id="4b8d2-167">将应用注册的所有实例 `YOUR_APP_ID_HERE` 替换为应用程序 ID。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-167">Replace all instances of `YOUR_APP_ID_HERE` with the application ID from your app registration.</span></span>

## <a name="configure-the-sample"></a><span data-ttu-id="4b8d2-168">配置示例</span><span class="sxs-lookup"><span data-stu-id="4b8d2-168">Configure the sample</span></span>

1. <span data-ttu-id="4b8d2-169">将文件 `example.env` 重命名为 `.env` 。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-169">Rename the `example.env` file to `.env`.</span></span>
1. <span data-ttu-id="4b8d2-170">编辑 `.env` 文件并做出以下更改。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-170">Edit the `.env` file and make the following changes.</span></span>
    1. <span data-ttu-id="4b8d2-171">替换为 `YOUR_APP_ID_HERE` 从 **应用** 注册门户获得的应用程序 ID。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-171">Replace `YOUR_APP_ID_HERE` with the **Application Id** you got from the App Registration Portal.</span></span>
    1. <span data-ttu-id="4b8d2-172">替换为 `YOUR_CLIENT_SECRET_HERE` 从应用注册门户获得的客户密码。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-172">Replace `YOUR_CLIENT_SECRET_HERE` with the client secret you got from the App Registration Portal.</span></span>
    1. <span data-ttu-id="4b8d2-173">替换为 `PATH_TO_LOCALHOST.CRT` 命令输出中 localhost.crt 文件 `npx office-addin-dev-certs install` 的路径。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-173">Replace `PATH_TO_LOCALHOST.CRT` with the path to your localhost.crt file from the output of the `npx office-addin-dev-certs install` command.</span></span>
    1. <span data-ttu-id="4b8d2-174">替换为 `PATH_TO_LOCALHOST.KEY` 命令输出中 localhost.key 文件 `npx office-addin-dev-certs install` 的路径。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-174">Replace `PATH_TO_LOCALHOST.KEY` with the path to your localhost.key file from the output of the `npx office-addin-dev-certs install` command.</span></span>

1. <span data-ttu-id="4b8d2-175">将文件 `config.example.js` 重命名为 `config.js` 。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-175">Rename the `config.example.js` file to `config.js`.</span></span>
1. <span data-ttu-id="4b8d2-176">编辑 `config.js` 文件并做出以下更改。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-176">Edit the `config.js` file and make the following changes.</span></span>
    1. <span data-ttu-id="4b8d2-177">替换为 `YOUR_APP_ID_HERE` 从 **应用** 注册门户获得的应用程序 ID。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-177">Replace `YOUR_APP_ID_HERE` with the **Application Id** you got from the App Registration Portal.</span></span>
1. <span data-ttu-id="4b8d2-178">在 CLI (命令行界面) ，导航到此目录并运行以下命令来安装要求。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-178">In your command-line interface (CLI), navigate to this directory and run the following command to install requirements.</span></span>

    ```Shell
    yarn install
    ```

## <a name="run-the-sample"></a><span data-ttu-id="4b8d2-179">运行示例</span><span class="sxs-lookup"><span data-stu-id="4b8d2-179">Run the sample</span></span>

1. <span data-ttu-id="4b8d2-180">在 CLI 中运行以下命令以启动应用程序。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-180">Run the following command in your CLI to start the application.</span></span>

    ```Shell
    yarn start
    ```

1. <span data-ttu-id="4b8d2-181">在浏览器中 [，转到Office.com](https://www.office.com/) 并登录。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-181">In your browser, go to [Office.com](https://www.office.com/) and sign in.</span></span> <span data-ttu-id="4b8d2-182">选择 **左侧** 工具栏中的"创建"，然后选择 **"电子表格"。**</span><span class="sxs-lookup"><span data-stu-id="4b8d2-182">Select **Create** in the left-hand toolbar, then select **Spreadsheet**.</span></span>

    !["创建"菜单的屏幕截图Office.com](/tutorial/images/office-select-excel.png)

1. <span data-ttu-id="4b8d2-184">选择 **"插入**"选项卡，然后选择 **"Office 加载项"。**</span><span class="sxs-lookup"><span data-stu-id="4b8d2-184">Select the **Insert** tab, then select **Office Add-ins**.</span></span>

1. <span data-ttu-id="4b8d2-185">选择 **"上载我的外接程序"，** 然后选择"**浏览"。**</span><span class="sxs-lookup"><span data-stu-id="4b8d2-185">Select **Upload My Add-in**, then select **Browse**.</span></span> <span data-ttu-id="4b8d2-186">上传manifest.xml **文件** 。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-186">Upload your **manifest.xml** file.</span></span>

1. <span data-ttu-id="4b8d2-187">选择 **"主页"\*\*\*\*选项卡上的"** 导入日历"按钮以打开任务窗格。</span><span class="sxs-lookup"><span data-stu-id="4b8d2-187">Select the **Import Calendar** button on the **Home** tab to open the taskpane.</span></span>

    !["开始"选项卡上"导入日历"按钮的屏幕截图](/tutorial/images/get-started.png)
