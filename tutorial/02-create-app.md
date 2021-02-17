---
ms.openlocfilehash: 70df2dfceb1d2df527acfecab894b28019c6febb
ms.sourcegitcommit: 24bb4b3df6a035806a58b609e1d8078ac5505fa1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/17/2021
ms.locfileid: "50274235"
---
<!-- markdownlint-disable MD002 MD041 -->

<span data-ttu-id="682bc-101">在此练习中，你将使用 Express 创建 Office 外接程序 [解决方案](http://expressjs.com/)。</span><span class="sxs-lookup"><span data-stu-id="682bc-101">In this exercise you will create an Office Add-in solution using [Express](http://expressjs.com/).</span></span> <span data-ttu-id="682bc-102">解决方案将包含两个部分。</span><span class="sxs-lookup"><span data-stu-id="682bc-102">The solution will consist of two parts.</span></span>

- <span data-ttu-id="682bc-103">加载项，作为静态 HTML 和 JavaScript 文件实现。</span><span class="sxs-lookup"><span data-stu-id="682bc-103">The add-in, implemented as static HTML and JavaScript files.</span></span>
- <span data-ttu-id="682bc-104">一Node.js/Express 服务器，该服务器为外接程序提供服务，并实现了一个 Web API 来检索加载项的数据。</span><span class="sxs-lookup"><span data-stu-id="682bc-104">A Node.js/Express server that serves the add-in and implements a web API to retrieve data for the add-in.</span></span>

## <a name="create-the-server"></a><span data-ttu-id="682bc-105">创建服务器</span><span class="sxs-lookup"><span data-stu-id="682bc-105">Create the server</span></span>

1. <span data-ttu-id="682bc-106">打开 CLI (命令行界面) 导航到要创建项目的目录，然后运行以下命令以生成package.js文件。</span><span class="sxs-lookup"><span data-stu-id="682bc-106">Open your command-line interface (CLI), navigate to a directory where you want to create your project, and run the following command to generate a package.json file.</span></span>

    ```Shell
    yarn init
    ```

    <span data-ttu-id="682bc-107">输入相应提示的值。</span><span class="sxs-lookup"><span data-stu-id="682bc-107">Enter values for the prompts as appropriate.</span></span> <span data-ttu-id="682bc-108">如果你不确定，默认值可以。</span><span class="sxs-lookup"><span data-stu-id="682bc-108">If you're unsure, the default values are fine.</span></span>

1. <span data-ttu-id="682bc-109">运行以下命令以安装依赖项。</span><span class="sxs-lookup"><span data-stu-id="682bc-109">Run the following commands to install dependencies.</span></span>

    ```Shell
    yarn add express@4.17.1 express-promise-router@4.0.1 dotenv@8.2.0 node-fetch@2.6.1 jsonwebtoken@8.5.1@
    yarn add jwks-rsa@1.11.0 @azure/msal-node@1.0.0-beta.1 @microsoft/microsoft-graph-client@2.1.1
    yarn add date-fns@2.16.1 date-fns-tz@1.0.12 isomorphic-fetch@3.0.0 windows-iana@4.2.1
    yarn add -D typescript@4.0.5 ts-node@9.0.0 nodemon@2.0.6 @types/node@14.14.7 @types/express@4.17.9
    yarn add -D @types/node-fetch@2.5.7 @types/jsonwebtoken@8.5.0 @types/microsoft-graph@1.26.0
    yarn add -D @types/office-js@1.0.147 @types/jquery@3.5.4 @types/isomorphic-fetch@0.0.35
    ```

1. <span data-ttu-id="682bc-110">运行以下命令以生成tsconfig.js文件。</span><span class="sxs-lookup"><span data-stu-id="682bc-110">Run the following command to generate a tsconfig.json file.</span></span>

    ```Shell
    tsc --init
    ```

1. <span data-ttu-id="682bc-111">在 **文本tsconfig.js打开 ./tsconfig.js，** 然后进行以下更改。</span><span class="sxs-lookup"><span data-stu-id="682bc-111">Open **./tsconfig.json** in a text editor and make the following changes.</span></span>

    - <span data-ttu-id="682bc-112">将 `target` 值更改为 `es6` 。</span><span class="sxs-lookup"><span data-stu-id="682bc-112">Change the `target` value to `es6`.</span></span>
    - <span data-ttu-id="682bc-113">取消注释 `outDir` 值，并设置为 `./dist` 。</span><span class="sxs-lookup"><span data-stu-id="682bc-113">Uncomment the `outDir` value and set it to `./dist`.</span></span>
    - <span data-ttu-id="682bc-114">取消注释 `rootDir` 值，并设置为 `./src` 。</span><span class="sxs-lookup"><span data-stu-id="682bc-114">Uncomment the `rootDir` value and set it to `./src`.</span></span>

1. <span data-ttu-id="682bc-115">打开 **./package.js，** 然后向 JSON 添加以下属性。</span><span class="sxs-lookup"><span data-stu-id="682bc-115">Open **./package.json** and add the following property to the JSON.</span></span>

    ```json
    "scripts": {
      "start": "nodemon ./src/server.ts",
      "build": "tsc --project ./"
    },
    ```

1. <span data-ttu-id="682bc-116">运行以下命令，为加载项生成和安装开发证书。</span><span class="sxs-lookup"><span data-stu-id="682bc-116">Run the following command to generate and install development certificates for your add-in.</span></span>

    ```Shell
    npx office-addin-dev-certs install
    ```

    <span data-ttu-id="682bc-117">如果系统提示确认，请确认操作。</span><span class="sxs-lookup"><span data-stu-id="682bc-117">If prompted for confirmation, confirm the actions.</span></span> <span data-ttu-id="682bc-118">命令完成后，你将看到如下所示的输出。</span><span class="sxs-lookup"><span data-stu-id="682bc-118">Once the command completes, you will see output similar to the following.</span></span>

    ```Shell
    You now have trusted access to https://localhost.
    Certificate: <path>\localhost.crt
    Key: <path>\localhost.key
    ```

1. <span data-ttu-id="682bc-119">在项目的根目录下创建一个名为 **.env** 的新文件，并添加以下代码。</span><span class="sxs-lookup"><span data-stu-id="682bc-119">Create a new file named **.env** in the root of your project and add the following code.</span></span>

    :::code language="ini" source="../demo/graph-tutorial/example.env":::

    <span data-ttu-id="682bc-120">用 `PATH_TO_LOCALHOST.CRT` localhost.crt 的路径和上一个命令 `PATH_TO_LOCALHOST.KEY` 指向 localhost.key 输出的路径替换。</span><span class="sxs-lookup"><span data-stu-id="682bc-120">Replace `PATH_TO_LOCALHOST.CRT` with the path to localhost.crt and `PATH_TO_LOCALHOST.KEY` with the path to localhost.key output by the previous command.</span></span>

1. <span data-ttu-id="682bc-121">在名为 **src** 的项目的根目录中创建新目录。</span><span class="sxs-lookup"><span data-stu-id="682bc-121">Create a new directory in the root of your project named **src**.</span></span>

1. <span data-ttu-id="682bc-122">在 **./src** 目录中创建两个目录 **：addin 和** **api。**</span><span class="sxs-lookup"><span data-stu-id="682bc-122">Create two directories in the **./src** directory: **addin** and **api**.</span></span>

1. <span data-ttu-id="682bc-123">在 **./src/api** 目录中创建一个名为 **auth.ts** 的新文件并添加以下代码。</span><span class="sxs-lookup"><span data-stu-id="682bc-123">Create a new file named **auth.ts** in the **./src/api** directory and add the following code.</span></span>

    ```typescript
    import Router from 'express-promise-router';

    const authRouter = Router();

    // TODO: Implement this router

    export default authRouter;
    ```

1. <span data-ttu-id="682bc-124">在 **./src/api** 目录中创建一个名为 **graph.ts** 的新文件并添加以下代码。</span><span class="sxs-lookup"><span data-stu-id="682bc-124">Create a new file named **graph.ts** in the **./src/api** directory and add the following code.</span></span>

    ```typescript
    import Router from 'express-promise-router';

    const graphRouter = Router();

    // TODO: Implement this router

    export default graphRouter;
    ```

1. <span data-ttu-id="682bc-125">在 **./src** 目录中创建一个名为 **server.ts** 的新文件，并添加以下代码。</span><span class="sxs-lookup"><span data-stu-id="682bc-125">Create a new file named **server.ts** in the **./src** directory and add the following code.</span></span>

    :::code language="typescript" source="../demo/graph-tutorial/src/server.ts" id="ServerSnippet":::

## <a name="create-the-add-in"></a><span data-ttu-id="682bc-126">创建加载项</span><span class="sxs-lookup"><span data-stu-id="682bc-126">Create the add-in</span></span>

1. <span data-ttu-id="682bc-127">在 **./src/addin** 目录中创建一 **个名为taskpane.html** 的新文件，并添加以下代码。</span><span class="sxs-lookup"><span data-stu-id="682bc-127">Create a new file named **taskpane.html** in the **./src/addin** directory and add the following code.</span></span>

    :::code language="html" source="../demo/graph-tutorial/src/addin/taskpane.html" id="TaskPaneHtmlSnippet":::

1. <span data-ttu-id="682bc-128">在 **./src/addin** 目录中创建一个名为 **taskpane.css** 的新文件，并添加以下代码。</span><span class="sxs-lookup"><span data-stu-id="682bc-128">Create a new file named **taskpane.css** in the **./src/addin** directory and add the following code.</span></span>

    :::code language="css" source="../demo/graph-tutorial/src/addin/taskpane.css":::

1. <span data-ttu-id="682bc-129">在 **./src/addin** **taskpane.js** 创建一个名为taskpane.js的新文件，并添加以下代码。</span><span class="sxs-lookup"><span data-stu-id="682bc-129">Create a new file named **taskpane.js** in the **./src/addin** directory and add the following code.</span></span>

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

1. <span data-ttu-id="682bc-130">在名为 assets 的 **.src/addin** 目录中创建新 **目录**。</span><span class="sxs-lookup"><span data-stu-id="682bc-130">Create a new directory in the **.src/addin** directory named **assets**.</span></span>

1. <span data-ttu-id="682bc-131">根据下表在此目录中添加三个 PNG 文件。</span><span class="sxs-lookup"><span data-stu-id="682bc-131">Add three PNG files in this directory according to the following table.</span></span>

    | <span data-ttu-id="682bc-132">文件名</span><span class="sxs-lookup"><span data-stu-id="682bc-132">File name</span></span>   | <span data-ttu-id="682bc-133">大小（以像素为单位）</span><span class="sxs-lookup"><span data-stu-id="682bc-133">Size in pixels</span></span> |
    |-------------|----------------|
    | <span data-ttu-id="682bc-134">icon-80.png</span><span class="sxs-lookup"><span data-stu-id="682bc-134">icon-80.png</span></span> | <span data-ttu-id="682bc-135">80x80</span><span class="sxs-lookup"><span data-stu-id="682bc-135">80x80</span></span>          |
    | <span data-ttu-id="682bc-136">icon-32.png</span><span class="sxs-lookup"><span data-stu-id="682bc-136">icon-32.png</span></span> | <span data-ttu-id="682bc-137">32x32</span><span class="sxs-lookup"><span data-stu-id="682bc-137">32x32</span></span>          |
    | <span data-ttu-id="682bc-138">icon-16.png</span><span class="sxs-lookup"><span data-stu-id="682bc-138">icon-16.png</span></span> | <span data-ttu-id="682bc-139">16x16</span><span class="sxs-lookup"><span data-stu-id="682bc-139">16x16</span></span>          |

    > [!NOTE]
    > <span data-ttu-id="682bc-140">可以使用此步骤需要的任何图像。</span><span class="sxs-lookup"><span data-stu-id="682bc-140">You can use any image you want for this step.</span></span> <span data-ttu-id="682bc-141">也可以直接从 [GitHub](https://github.com/microsoftgraph/msgraph-training-office-addin/demo/graph-tutorial/src/addin/assets)下载此示例中使用的图像。</span><span class="sxs-lookup"><span data-stu-id="682bc-141">You can also download the images used in this sample directly from [GitHub](https://github.com/microsoftgraph/msgraph-training-office-addin/demo/graph-tutorial/src/addin/assets).</span></span>

1. <span data-ttu-id="682bc-142">在名为清单的项目的根目录中创建新 **目录**。</span><span class="sxs-lookup"><span data-stu-id="682bc-142">Create a new directory in the root of the project named **manifest**.</span></span>

1. <span data-ttu-id="682bc-143">在 ./manifest **manifest.xml** 创建一个名为 **manifest.xml** 的新文件，并添加以下代码。</span><span class="sxs-lookup"><span data-stu-id="682bc-143">Create a new file named **manifest.xml** in the **./manifest** folder and add the following code.</span></span> <span data-ttu-id="682bc-144">替换为 `NEW_GUID_HERE` 新的 GUID，如 `b4fa03b8-1eb6-4e8b-a380-e0476be9e019` 。</span><span class="sxs-lookup"><span data-stu-id="682bc-144">Replace `NEW_GUID_HERE` with a new GUID, like `b4fa03b8-1eb6-4e8b-a380-e0476be9e019`.</span></span>

    :::code language="xml" source="../demo/graph-tutorial/manifest/manifest.xml":::

## <a name="side-load-the-add-in-in-excel"></a><span data-ttu-id="682bc-145">在 Excel 中旁加载加载项</span><span class="sxs-lookup"><span data-stu-id="682bc-145">Side-load the add-in in Excel</span></span>

1. <span data-ttu-id="682bc-146">通过运行以下命令启动服务器。</span><span class="sxs-lookup"><span data-stu-id="682bc-146">Start the server by running the following command.</span></span>

    ```Shell
    yarn start
    ```

1. <span data-ttu-id="682bc-147">打开浏览器并浏览到 `https://localhost:3000/taskpane.html` 。</span><span class="sxs-lookup"><span data-stu-id="682bc-147">Open your browser and browse to `https://localhost:3000/taskpane.html`.</span></span> <span data-ttu-id="682bc-148">你应该会看到一 `Not loaded` 条消息。</span><span class="sxs-lookup"><span data-stu-id="682bc-148">You should see a `Not loaded` message.</span></span>

1. <span data-ttu-id="682bc-149">在浏览器中 [，转到Office.com](https://www.office.com/) 并登录。</span><span class="sxs-lookup"><span data-stu-id="682bc-149">In your browser, go to [Office.com](https://www.office.com/) and sign in.</span></span> <span data-ttu-id="682bc-150">选择 **左侧** 工具栏中的"创建"，然后选择 **"电子表格"。**</span><span class="sxs-lookup"><span data-stu-id="682bc-150">Select **Create** in the left-hand toolbar, then select **Spreadsheet**.</span></span>

    !["创建"菜单的屏幕截图Office.com](images/office-select-excel.png)

1. <span data-ttu-id="682bc-152">选择 **"插入**"选项卡，然后选择 **"Office 加载项"。**</span><span class="sxs-lookup"><span data-stu-id="682bc-152">Select the **Insert** tab, then select **Office Add-ins**.</span></span>

1. <span data-ttu-id="682bc-153">选择 **"上载我的外接程序"，** 然后选择"**浏览"。**</span><span class="sxs-lookup"><span data-stu-id="682bc-153">Select **Upload My Add-in**, then select **Browse**.</span></span> <span data-ttu-id="682bc-154">上传 **./manifest/manifest.xml** 文件。</span><span class="sxs-lookup"><span data-stu-id="682bc-154">Upload your **./manifest/manifest.xml** file.</span></span>

1. <span data-ttu-id="682bc-155">选择 **"主页"\*\*\*\*选项卡上的"** 导入日历"按钮以打开任务窗格。</span><span class="sxs-lookup"><span data-stu-id="682bc-155">Select the **Import Calendar** button on the **Home** tab to open the taskpane.</span></span>

    !["开始"选项卡上"导入日历"按钮的屏幕截图](images/get-started.png)

1. <span data-ttu-id="682bc-157">任务窗格打开后，应该会看到 `Hello World!` 一条消息。</span><span class="sxs-lookup"><span data-stu-id="682bc-157">After the taskpane opens, you should see a `Hello World!` message.</span></span>

    ![Hello World 消息的屏幕截图](images/hello-world.png)
