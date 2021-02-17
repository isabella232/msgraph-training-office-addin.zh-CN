---
ms.openlocfilehash: 69d77c19cb2c589086df2b4bf19eeadf41aad58a
ms.sourcegitcommit: 24bb4b3df6a035806a58b609e1d8078ac5505fa1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/17/2021
ms.locfileid: "50274238"
---
<!-- markdownlint-disable MD002 MD041 -->

在此练习中，你将在外接程序中启用 Office 外接程序单一登录 [ (SSO) ， ](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins) 并扩展 Web API 以支持代表 [流](https://docs.microsoft.com/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow)。 这是获取调用 Microsoft Graph 所需的 OAuth 访问令牌所必需的。

## <a name="overview"></a>概述

Office 外接程序 SSO 提供了访问令牌，但该令牌仅允许外接程序调用它自己的 Web API。 它不支持直接访问 Microsoft Graph。 此过程的工作方式如下。

1. 加载项通过调用 [getAccessToken](https://docs.microsoft.com/javascript/api/office-runtime/officeruntime.auth?view=common-js#getaccesstoken-options-)获取令牌。 声明令牌的 (是) 应用程序注册的应用程序 `aud` ID。
1. 加载项在调用 Web API 时，在标头 `Authorization` 中发送此令牌。
1. Web API 验证令牌，然后使用代表流来交换此令牌以交换 Microsoft Graph 令牌。 此新令牌的受众为 `https://graph.microsoft.com` 。
1. Web API 使用新令牌调用 Microsoft Graph，将结果返回给加载项。

## <a name="configure-the-solution"></a>配置解决方案

1. 打开 **./.env，** 并更新应用注册中的应用程序 `AZURE_APP_ID` ID `AZURE_CLIENT_SECRET` 和客户端密码。

    > [!IMPORTANT]
    > 如果你使用的是源代码管理（如 git），那么现在应该将 **.env** 文件从源代码管理中排除，以避免意外泄露应用 ID 和客户端密码。

1. 打开 **./manifest/manifest.xml，** 并替换应用注册中的应用程序 `YOUR_APP_ID_HERE` ID 的所有实例。

1. 在名为config.js的 **./src/addin** 目录中创建新文件，并添加以下代码，替换为应用注册 `YOUR_APP_ID_HERE` 中的应用程序 ID。

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/config.example.js":::

## <a name="implement-sign-in"></a>实现登录

1. 打开 **./src/api/auth.ts，** 在文件顶部 `import` 添加以下语句。

    ```typescript
    import jwt, { SigningKeyCallback, JwtHeader } from 'jsonwebtoken';
    import jwksClient from 'jwks-rsa';
    import * as msal from '@azure/msal-node';
    ```

1. 在语句后添加以下 `import` 代码。

    :::code language="typescript" source="../demo/graph-tutorial/src/api/auth.ts" id="TokenExchangeSnippet":::

    此代码 [初始化 MSAL](https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-node/docs/initialize-confidential-client-application.md)机密客户端，并导出函数以从加载项发送的令牌获取 Graph 令牌。

1. 在行前添加以下 `export default authRouter;` 代码。

    :::code language="typescript" source="../demo/graph-tutorial/src/api/auth.ts" id="GetAuthStatusSnippet":::

    此代码实现 API () ，用于检查加载项令牌能否以无提示方式交换 `GET /auth/status` Graph 令牌。 加载项将使用此 API 确定是否需要向用户显示交互式登录名。

1. 打开 **./src/addin/taskpane.js，** 然后向文件添加以下代码。

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="AuthUiSnippet":::

    此代码添加函数以更新 UI，并使用 Office [对话框 API](https://docs.microsoft.com/office/dev/add-ins/develop/dialog-api-in-office-add-ins) 启动交互式身份验证流。

1. 添加以下函数来实现临时主 UI。

    ```javascript
    function showMainUi() {
      $('.container').empty();
      $('<p/>', {
        class: 'ms-fontSize-24 ms-fontWeight-bold',
        text: 'Authenticated!'
      }).appendTo('.container');
    }
    ```

1. 将现有 `Office.onReady` 调用替换为以下内容。

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="OfficeReadySnippet":::

    考虑此代码执行什么功能。

    - 首次加载任务窗格时，它将调用获取一个范围为加载项 `getAccessToken` Web API 的令牌。
    - 它使用该令牌调用 API，以检查用户是否同意 `/auth/status` Microsoft Graph 范围。
        - 如果用户尚未同意，它使用弹出窗口通过交互式登录征得用户同意。
        - 如果用户已同意，它将加载主 UI。

### <a name="getting-user-consent"></a>获取用户同意

即使外接程序使用的是 SSO，用户仍必须同意加载项通过 Microsoft Graph 访问其数据。 获取同意是一个一次过程。 用户授予同意后，SSO 令牌可以交换为 Graph 令牌，无需任何用户交互。 在此部分中，你将使用 [msal-browser](https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/lib/msal-browser)在外接程序中实现同意体验。

1. 在名为consent.js的 **./src/addin** **目录中** 创建新文件，并添加以下代码。

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/consent.js" id="ConsentJsSnippet":::

    此代码为用户登录，请求在应用注册时配置的 Microsoft Graph 权限集。

1. 在名为consent.html 的 **./src/addin** **目录中** 创建新文件并添加以下代码。

    :::code language="html" source="../demo/graph-tutorial/src/addin/consent.html" id="ConsentHtmlSnippet":::

    此代码实现一个基本 HTML 页面，以 **加载consent.js文件** 。 此页面将在弹出对话框中加载。

1. 保存所有更改并重新启动服务器。

1. 在 Excel **中manifest.xml** 加载外接程序时，使用相同的步骤重新 [上传您的加载项文件](02-create-app.md#side-load-the-add-in-in-excel)。

1. 选择 **"主页"** 选项卡上的"导入日历 **"** 按钮以打开任务窗格。

1. 选择 **任务窗格中** 的"授予权限"按钮以在弹出窗口中启动同意对话框。 登录并授予同意。

1. 任务窗格更新为"Authenticated！" 消息。 可以按如下方式检查令牌。

    - 在浏览器的开发人员工具中，API 令牌显示在控制台中。
    - 在运行应用程序服务器的 CLI 中，Node.js Graph 令牌。

    可以在以下时间比较这些标记 [https://jwt.ms](https://jwt.ms) 。 请注意，API 令牌的受众 () 设置为应用注册的应用程序 ID，并且 () `aud` `scp` 为 `access_as_user` 。
