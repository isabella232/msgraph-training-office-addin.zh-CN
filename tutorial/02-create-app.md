---
ms.openlocfilehash: 09ef719985094d87c438b6f98b931865dbd59c75
ms.sourcegitcommit: 8a65c826f6b229c287a782d784b6d9629aa5a3d0
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/20/2021
ms.locfileid: "51899420"
---
<!-- markdownlint-disable MD002 MD041 -->

在此练习中，你将使用 Express 创建 Office 外接程序 [解决方案](http://expressjs.com/)。 解决方案由两部分组成。

- 加载项，作为静态 HTML 和 JavaScript 文件实现。
- A Node.js/Express server that serves the add-in and implements a web API to retrieve data for the add-in.

## <a name="create-the-server"></a>创建服务器

1. 使用 CLI (打开) 界面，导航到要创建项目的目录，然后运行以下命令以生成package.js文件。

    ```Shell
    yarn init
    ```

    根据情况输入提示值。 如果你不确定，默认值可以。

1. 运行以下命令以安装依赖项。

    ```Shell
    yarn add express@4.17.1 express-promise-router@4.1.0 dotenv@8.2.0 node-fetch@2.6.1 jsonwebtoken@8.5.1@
    yarn add jwks-rsa@2.0.2 @azure/msal-node@1.0.2 @microsoft/microsoft-graph-client@2.2.1
    yarn add date-fns@2.21.1 date-fns-tz@1.1.4 isomorphic-fetch@3.0.0 windows-iana@5.0.1
    yarn add -D typescript@4.2.4 ts-node@9.1.1 nodemon@2.0.7 @types/node@14.14.41 @types/express@4.17.11
    yarn add -D @types/node-fetch@2.5.10 @types/jsonwebtoken@8.5.1 @types/microsoft-graph@1.35.0
    yarn add -D @types/office-js@1.0.174 @types/jquery@3.5.5 @types/isomorphic-fetch@0.0.35
    ```

1. 运行以下命令以生成tsconfig.js文件。

    ```Shell
    tsc --init
    ```

1. 在 **文本tsconfig.js打开 "./tsconfig.js** on"，然后进行以下更改。

    - 将 `target` 值更改为 `es6` 。
    - 取消注释 `outDir` 值，并设置为 `./dist` 。
    - 取消注释 `rootDir` 值，并设置为 `./src` 。

1. 打开 **"package.js"，** 然后向 JSON 添加以下属性。

    ```json
    "scripts": {
      "start": "nodemon ./src/server.ts",
      "build": "tsc --project ./"
    },
    ```

1. 运行以下命令，为加载项生成和安装开发证书。

    ```Shell
    npx office-addin-dev-certs install
    ```

    如果系统提示确认，请确认操作。 命令完成后，你将看到与以下内容类似的输出。

    ```Shell
    You now have trusted access to https://localhost.
    Certificate: <path>\localhost.crt
    Key: <path>\localhost.key
    ```

1. 在项目的根目录下创建一个名为 **.env** 的新文件，并添加以下代码。

    :::code language="ini" source="../demo/graph-tutorial/example.env":::

    将 替换为 localhost.crt 的路径和上一个命令指向 `PATH_TO_LOCALHOST.CRT` `PATH_TO_LOCALHOST.KEY` localhost.key 输出的路径。

1. 在名为 **src** 的项目的根目录中新建目录。

1. 在 **./src** 目录中创建两个目录 **：addin 和** **api**。

1. 在 **./src/api** 目录中创建一个名为 **auth.ts** 的新文件，并添加以下代码。

    ```typescript
    import Router from 'express-promise-router';

    const authRouter = Router();

    // TODO: Implement this router

    export default authRouter;
    ```

1. 在 **./src/api** 目录中创建一个名为 **graph.ts** 的新文件，并添加以下代码。

    ```typescript
    import Router from 'express-promise-router';

    const graphRouter = Router();

    // TODO: Implement this router

    export default graphRouter;
    ```

1. 在 **./src** 目录中创建一个名为 **server.ts** 的新文件，并添加以下代码。

    :::code language="typescript" source="../demo/graph-tutorial/src/server.ts" id="ServerSnippet":::

## <a name="create-the-add-in"></a>创建加载项

1. 在 **./src/addin** taskpane.htm创建一个名为taskpane.htm **l** 的新文件，并添加以下代码。

    :::code language="html" source="../demo/graph-tutorial/src/addin/taskpane.html" id="TaskPaneHtmlSnippet":::

1. 在 **./src/addin** 目录中创建一个名为 **taskpane.css** 的新文件，并添加以下代码。

    :::code language="css" source="../demo/graph-tutorial/src/addin/taskpane.css":::

1. 在 **./src/addin** **taskpane.js** 创建一个名为"addin"的新文件，并添加以下代码。

    ```javascript
    // TEMPORARY CODE TO VERIFY ADD-IN LOADS
    'use strict';

    Office.onReady(info => {
      if (info.host === Office.HostType.Excel) {
        $(function() {
          $('p').text('Hello World!!');
        });
      }
    });
    ```

1. 在名为 assets 的 **.src/addin** 目录中创建新 **目录**。

1. 根据下表，在此目录中添加三个 PNG 文件。

    | 文件名   | 大小（以像素为单位） |
    |-------------|----------------|
    | icon-80.png | 80x80          |
    | icon-32.png | 32x32          |
    | icon-16.png | 16x16          |

    > [!NOTE]
    > 可以使用此步骤需要的任何图像。 您还可以直接从 [GitHub](https://github.com/microsoftgraph/msgraph-training-office-addin/demo/graph-tutorial/src/addin/assets)下载此示例中使用的图像。

1. 在名为 清单 的项目的根目录中创建新 **目录**。

1. 在 **./manifest****文件夹中manifest.xml** 一个名为manifest.xml文件，并添加以下代码。 将 `NEW_GUID_HERE` 替换为新的 GUID，如 `b4fa03b8-1eb6-4e8b-a380-e0476be9e019` 。

    :::code language="xml" source="../demo/graph-tutorial/manifest/manifest.xml":::

## <a name="side-load-the-add-in-in-excel"></a>在 Excel 中旁加载加载项

1. 通过运行以下命令启动服务器。

    ```Shell
    yarn start
    ```

1. 打开浏览器并浏览到 `https://localhost:3000/taskpane.html` 。 你应该会看到一 `Not loaded` 条消息。

1. 在浏览器中，转到 [Office.com](https://www.office.com/) 并登录。 选择 **左侧** 工具栏中的"创建"，然后选择"电子表格 **"。**

    !["创建"菜单的屏幕截图 Office.com](images/office-select-excel.png)

1. 选择"**插入**"选项卡，然后选择 **"Office 外接程序"。**

1. 选择 **"上载我的外接程序"，** 然后选择"浏览 **"。** 上传 **./manifest/manifest.xml** 文件。

1. 选择" **开始"选项卡** 上的"导入日历 **"** 按钮以打开任务窗格。

    !["开始"选项卡上"导入日历"按钮的屏幕截图](images/get-started.png)

1. 任务窗格打开后，应该会看到一 `Hello World!` 条消息。

    ![Hello World 消息的屏幕截图](images/hello-world.png)
