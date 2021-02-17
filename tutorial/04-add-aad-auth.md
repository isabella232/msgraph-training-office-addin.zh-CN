---
ms.openlocfilehash: 69d77c19cb2c589086df2b4bf19eeadf41aad58a
ms.sourcegitcommit: 24bb4b3df6a035806a58b609e1d8078ac5505fa1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/17/2021
ms.locfileid: "50274238"
---
<!-- markdownlint-disable MD002 MD041 -->

<span data-ttu-id="53ad5-101">在此练习中，你将在外接程序中启用 Office 外接程序单一登录 [ (SSO) ， ](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins) 并扩展 Web API 以支持代表 [流](https://docs.microsoft.com/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow)。</span><span class="sxs-lookup"><span data-stu-id="53ad5-101">In this exercise you will enable [Office Add-in single sign-on (SSO)](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins) in the add-in, and extend the web API to support [on-behalf-of flow](https://docs.microsoft.com/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow).</span></span> <span data-ttu-id="53ad5-102">这是获取调用 Microsoft Graph 所需的 OAuth 访问令牌所必需的。</span><span class="sxs-lookup"><span data-stu-id="53ad5-102">This is required to obtain the necessary OAuth access token to call the Microsoft Graph.</span></span>

## <a name="overview"></a><span data-ttu-id="53ad5-103">概述</span><span class="sxs-lookup"><span data-stu-id="53ad5-103">Overview</span></span>

<span data-ttu-id="53ad5-104">Office 外接程序 SSO 提供了访问令牌，但该令牌仅允许外接程序调用它自己的 Web API。</span><span class="sxs-lookup"><span data-stu-id="53ad5-104">Office Add-in SSO provides an access token, but that token is only enables the add-in to call it's own web API.</span></span> <span data-ttu-id="53ad5-105">它不支持直接访问 Microsoft Graph。</span><span class="sxs-lookup"><span data-stu-id="53ad5-105">It does not enable direct access to the Microsoft Graph.</span></span> <span data-ttu-id="53ad5-106">此过程的工作方式如下。</span><span class="sxs-lookup"><span data-stu-id="53ad5-106">The process works as follows.</span></span>

1. <span data-ttu-id="53ad5-107">加载项通过调用 [getAccessToken](https://docs.microsoft.com/javascript/api/office-runtime/officeruntime.auth?view=common-js#getaccesstoken-options-)获取令牌。</span><span class="sxs-lookup"><span data-stu-id="53ad5-107">The add-in gets a token by calling [getAccessToken](https://docs.microsoft.com/javascript/api/office-runtime/officeruntime.auth?view=common-js#getaccesstoken-options-).</span></span> <span data-ttu-id="53ad5-108">声明令牌的 (是) 应用程序注册的应用程序 `aud` ID。</span><span class="sxs-lookup"><span data-stu-id="53ad5-108">This token's audience (the `aud` claim) is the application ID of the add-in's app registration.</span></span>
1. <span data-ttu-id="53ad5-109">加载项在调用 Web API 时，在标头 `Authorization` 中发送此令牌。</span><span class="sxs-lookup"><span data-stu-id="53ad5-109">The add-in sends this token in the `Authorization` header when it makes a call to the web API.</span></span>
1. <span data-ttu-id="53ad5-110">Web API 验证令牌，然后使用代表流来交换此令牌以交换 Microsoft Graph 令牌。</span><span class="sxs-lookup"><span data-stu-id="53ad5-110">The web API validates the token, then uses the on-behalf-of flow to exchange this token for a Microsoft Graph token.</span></span> <span data-ttu-id="53ad5-111">此新令牌的受众为 `https://graph.microsoft.com` 。</span><span class="sxs-lookup"><span data-stu-id="53ad5-111">This new token's audience is `https://graph.microsoft.com`.</span></span>
1. <span data-ttu-id="53ad5-112">Web API 使用新令牌调用 Microsoft Graph，将结果返回给加载项。</span><span class="sxs-lookup"><span data-stu-id="53ad5-112">The web API uses the new token to make calls to the Microsoft Graph, and returns the results back to the add-in.</span></span>

## <a name="configure-the-solution"></a><span data-ttu-id="53ad5-113">配置解决方案</span><span class="sxs-lookup"><span data-stu-id="53ad5-113">Configure the solution</span></span>

1. <span data-ttu-id="53ad5-114">打开 **./.env，** 并更新应用注册中的应用程序 `AZURE_APP_ID` ID `AZURE_CLIENT_SECRET` 和客户端密码。</span><span class="sxs-lookup"><span data-stu-id="53ad5-114">Open **./.env** and update the `AZURE_APP_ID` and `AZURE_CLIENT_SECRET` with the application ID and client secret from your app registration.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="53ad5-115">如果你使用的是源代码管理（如 git），那么现在应该将 **.env** 文件从源代码管理中排除，以避免意外泄露应用 ID 和客户端密码。</span><span class="sxs-lookup"><span data-stu-id="53ad5-115">If you're using source control such as git, now would be a good time to exclude the **.env** file from source control to avoid inadvertently leaking your app ID and client secret.</span></span>

1. <span data-ttu-id="53ad5-116">打开 **./manifest/manifest.xml，** 并替换应用注册中的应用程序 `YOUR_APP_ID_HERE` ID 的所有实例。</span><span class="sxs-lookup"><span data-stu-id="53ad5-116">Open **./manifest/manifest.xml** and replace all instances of `YOUR_APP_ID_HERE` with the application ID from your app registration.</span></span>

1. <span data-ttu-id="53ad5-117">在名为config.js的 **./src/addin** 目录中创建新文件，并添加以下代码，替换为应用注册 `YOUR_APP_ID_HERE` 中的应用程序 ID。</span><span class="sxs-lookup"><span data-stu-id="53ad5-117">Create a new file in the **./src/addin** directory named **config.js** and add the following code, replacing `YOUR_APP_ID_HERE` with the application ID from your app registration.</span></span>

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/config.example.js":::

## <a name="implement-sign-in"></a><span data-ttu-id="53ad5-118">实现登录</span><span class="sxs-lookup"><span data-stu-id="53ad5-118">Implement sign-in</span></span>

1. <span data-ttu-id="53ad5-119">打开 **./src/api/auth.ts，** 在文件顶部 `import` 添加以下语句。</span><span class="sxs-lookup"><span data-stu-id="53ad5-119">Open **./src/api/auth.ts** and add the following `import` statements at the top of the file.</span></span>

    ```typescript
    import jwt, { SigningKeyCallback, JwtHeader } from 'jsonwebtoken';
    import jwksClient from 'jwks-rsa';
    import * as msal from '@azure/msal-node';
    ```

1. <span data-ttu-id="53ad5-120">在语句后添加以下 `import` 代码。</span><span class="sxs-lookup"><span data-stu-id="53ad5-120">Add the following code after the `import` statements.</span></span>

    :::code language="typescript" source="../demo/graph-tutorial/src/api/auth.ts" id="TokenExchangeSnippet":::

    <span data-ttu-id="53ad5-121">此代码 [初始化 MSAL](https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-node/docs/initialize-confidential-client-application.md)机密客户端，并导出函数以从加载项发送的令牌获取 Graph 令牌。</span><span class="sxs-lookup"><span data-stu-id="53ad5-121">This code [initializes an MSAL confidential client](https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-node/docs/initialize-confidential-client-application.md), and exports a function to get a Graph token from the token sent by the add-in.</span></span>

1. <span data-ttu-id="53ad5-122">在行前添加以下 `export default authRouter;` 代码。</span><span class="sxs-lookup"><span data-stu-id="53ad5-122">Add the following code before the `export default authRouter;` line.</span></span>

    :::code language="typescript" source="../demo/graph-tutorial/src/api/auth.ts" id="GetAuthStatusSnippet":::

    <span data-ttu-id="53ad5-123">此代码实现 API () ，用于检查加载项令牌能否以无提示方式交换 `GET /auth/status` Graph 令牌。</span><span class="sxs-lookup"><span data-stu-id="53ad5-123">This code implements an API (`GET /auth/status`) that checks if the add-in token can be silently exchanged for a Graph token.</span></span> <span data-ttu-id="53ad5-124">加载项将使用此 API 确定是否需要向用户显示交互式登录名。</span><span class="sxs-lookup"><span data-stu-id="53ad5-124">The add-in will use this API to determine if it needs to present an interactive login to the user.</span></span>

1. <span data-ttu-id="53ad5-125">打开 **./src/addin/taskpane.js，** 然后向文件添加以下代码。</span><span class="sxs-lookup"><span data-stu-id="53ad5-125">Open **./src/addin/taskpane.js** and add the following code to the file.</span></span>

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="AuthUiSnippet":::

    <span data-ttu-id="53ad5-126">此代码添加函数以更新 UI，并使用 Office [对话框 API](https://docs.microsoft.com/office/dev/add-ins/develop/dialog-api-in-office-add-ins) 启动交互式身份验证流。</span><span class="sxs-lookup"><span data-stu-id="53ad5-126">This code adds functions to update the UI, and to use the [Office Dialog API](https://docs.microsoft.com/office/dev/add-ins/develop/dialog-api-in-office-add-ins) to initiate an interactive authentication flow.</span></span>

1. <span data-ttu-id="53ad5-127">添加以下函数来实现临时主 UI。</span><span class="sxs-lookup"><span data-stu-id="53ad5-127">Add the following function to implement a temporary main UI.</span></span>

    ```javascript
    function showMainUi() {
      $('.container').empty();
      $('<p/>', {
        class: 'ms-fontSize-24 ms-fontWeight-bold',
        text: 'Authenticated!'
      }).appendTo('.container');
    }
    ```

1. <span data-ttu-id="53ad5-128">将现有 `Office.onReady` 调用替换为以下内容。</span><span class="sxs-lookup"><span data-stu-id="53ad5-128">Replace the existing `Office.onReady` call with the following.</span></span>

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="OfficeReadySnippet":::

    <span data-ttu-id="53ad5-129">考虑此代码执行什么功能。</span><span class="sxs-lookup"><span data-stu-id="53ad5-129">Consider what this code does.</span></span>

    - <span data-ttu-id="53ad5-130">首次加载任务窗格时，它将调用获取一个范围为加载项 `getAccessToken` Web API 的令牌。</span><span class="sxs-lookup"><span data-stu-id="53ad5-130">When the task pane first loads, it calls `getAccessToken` to get a token scoped for the add-in's web API.</span></span>
    - <span data-ttu-id="53ad5-131">它使用该令牌调用 API，以检查用户是否同意 `/auth/status` Microsoft Graph 范围。</span><span class="sxs-lookup"><span data-stu-id="53ad5-131">It uses that token to call the `/auth/status` API to check if the user has given consent to the Microsoft Graph scopes yet.</span></span>
        - <span data-ttu-id="53ad5-132">如果用户尚未同意，它使用弹出窗口通过交互式登录征得用户同意。</span><span class="sxs-lookup"><span data-stu-id="53ad5-132">If the user has not consented, it uses a pop-up window to get the user's consent through an interactive login.</span></span>
        - <span data-ttu-id="53ad5-133">如果用户已同意，它将加载主 UI。</span><span class="sxs-lookup"><span data-stu-id="53ad5-133">If the user has consented, it loads the main UI.</span></span>

### <a name="getting-user-consent"></a><span data-ttu-id="53ad5-134">获取用户同意</span><span class="sxs-lookup"><span data-stu-id="53ad5-134">Getting user consent</span></span>

<span data-ttu-id="53ad5-135">即使外接程序使用的是 SSO，用户仍必须同意加载项通过 Microsoft Graph 访问其数据。</span><span class="sxs-lookup"><span data-stu-id="53ad5-135">Even though the add-in is using SSO, the user still has to consent to the add-in accessing their data via Microsoft Graph.</span></span> <span data-ttu-id="53ad5-136">获取同意是一个一次过程。</span><span class="sxs-lookup"><span data-stu-id="53ad5-136">Getting consent is a one-time process.</span></span> <span data-ttu-id="53ad5-137">用户授予同意后，SSO 令牌可以交换为 Graph 令牌，无需任何用户交互。</span><span class="sxs-lookup"><span data-stu-id="53ad5-137">Once the user has granted consent, the SSO token can be exchanged for a Graph token without any user interaction.</span></span> <span data-ttu-id="53ad5-138">在此部分中，你将使用 [msal-browser](https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/lib/msal-browser)在外接程序中实现同意体验。</span><span class="sxs-lookup"><span data-stu-id="53ad5-138">In this section you'll implement the consent experience in the add-in using [msal-browser](https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/lib/msal-browser).</span></span>

1. <span data-ttu-id="53ad5-139">在名为consent.js的 **./src/addin** **目录中** 创建新文件，并添加以下代码。</span><span class="sxs-lookup"><span data-stu-id="53ad5-139">Create a new file in the **./src/addin** directory named **consent.js** and add the following code.</span></span>

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/consent.js" id="ConsentJsSnippet":::

    <span data-ttu-id="53ad5-140">此代码为用户登录，请求在应用注册时配置的 Microsoft Graph 权限集。</span><span class="sxs-lookup"><span data-stu-id="53ad5-140">This code does login for the user, requesting the set of Microsoft Graph permissions that are configured on the app registration.</span></span>

1. <span data-ttu-id="53ad5-141">在名为consent.html 的 **./src/addin** **目录中** 创建新文件并添加以下代码。</span><span class="sxs-lookup"><span data-stu-id="53ad5-141">Create a new file in the **./src/addin** directory named **consent.html** and add the following code.</span></span>

    :::code language="html" source="../demo/graph-tutorial/src/addin/consent.html" id="ConsentHtmlSnippet":::

    <span data-ttu-id="53ad5-142">此代码实现一个基本 HTML 页面，以 **加载consent.js文件** 。</span><span class="sxs-lookup"><span data-stu-id="53ad5-142">This code implements a basic HTML page to load the **consent.js** file.</span></span> <span data-ttu-id="53ad5-143">此页面将在弹出对话框中加载。</span><span class="sxs-lookup"><span data-stu-id="53ad5-143">This page will be loaded in a pop-up dialog.</span></span>

1. <span data-ttu-id="53ad5-144">保存所有更改并重新启动服务器。</span><span class="sxs-lookup"><span data-stu-id="53ad5-144">Save all of your changes and restart the server.</span></span>

1. <span data-ttu-id="53ad5-145">在 Excel **中manifest.xml** 加载外接程序时，使用相同的步骤重新 [上传您的加载项文件](02-create-app.md#side-load-the-add-in-in-excel)。</span><span class="sxs-lookup"><span data-stu-id="53ad5-145">Re-upload your **manifest.xml** file using the same steps in [Side-load the add-in in Excel](02-create-app.md#side-load-the-add-in-in-excel).</span></span>

1. <span data-ttu-id="53ad5-146">选择 **"主页"** 选项卡上的"导入日历 **"** 按钮以打开任务窗格。</span><span class="sxs-lookup"><span data-stu-id="53ad5-146">Select the **Import Calendar** button on the **Home** tab to open the task pane.</span></span>

1. <span data-ttu-id="53ad5-147">选择 **任务窗格中** 的"授予权限"按钮以在弹出窗口中启动同意对话框。</span><span class="sxs-lookup"><span data-stu-id="53ad5-147">Select the **Give permission** button in the task pane to launch the consent dialog in a pop-up window.</span></span> <span data-ttu-id="53ad5-148">登录并授予同意。</span><span class="sxs-lookup"><span data-stu-id="53ad5-148">Sign in and grant consent.</span></span>

1. <span data-ttu-id="53ad5-149">任务窗格更新为"Authenticated！"</span><span class="sxs-lookup"><span data-stu-id="53ad5-149">The task pane updates with an "Authenticated!"</span></span> <span data-ttu-id="53ad5-150">消息。</span><span class="sxs-lookup"><span data-stu-id="53ad5-150">message.</span></span> <span data-ttu-id="53ad5-151">可以按如下方式检查令牌。</span><span class="sxs-lookup"><span data-stu-id="53ad5-151">You can check the tokens as follows.</span></span>

    - <span data-ttu-id="53ad5-152">在浏览器的开发人员工具中，API 令牌显示在控制台中。</span><span class="sxs-lookup"><span data-stu-id="53ad5-152">In your brower's developer tools, the API token is shown in the Console.</span></span>
    - <span data-ttu-id="53ad5-153">在运行应用程序服务器的 CLI 中，Node.js Graph 令牌。</span><span class="sxs-lookup"><span data-stu-id="53ad5-153">In your CLI where you are running the Node.js server, the Graph token is printed.</span></span>

    <span data-ttu-id="53ad5-154">可以在以下时间比较这些标记 [https://jwt.ms](https://jwt.ms) 。</span><span class="sxs-lookup"><span data-stu-id="53ad5-154">You can compare these token at [https://jwt.ms](https://jwt.ms).</span></span> <span data-ttu-id="53ad5-155">请注意，API 令牌的受众 () 设置为应用注册的应用程序 ID，并且 () `aud` `scp` 为 `access_as_user` 。</span><span class="sxs-lookup"><span data-stu-id="53ad5-155">Notice that the API token's audience (`aud`) is set to the application ID of your app registration, and the scope (`scp`) is `access_as_user`.</span></span>
