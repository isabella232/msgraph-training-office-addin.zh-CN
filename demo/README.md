---
ms.openlocfilehash: fed56663591ac36e4defcae3a8d0a222f86e3026
ms.sourcegitcommit: 24bb4b3df6a035806a58b609e1d8078ac5505fa1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/17/2021
ms.locfileid: "50274231"
---
# <a name="how-to-run-the-completed-project"></a>如何运行已完成的项目

## <a name="prerequisites"></a>先决条件

若要在此文件夹中运行已完成的项目，需要执行以下操作：

- [Node.js](https://nodejs.org) 计算机上安装的Node.js和 [一](https://yarnpkg.com/) 线。  (：本教程是使用 Node 版本 14.15.0 和一线 1.22.0 编写的。 本指南中的步骤可能与其他版本一起运行，但尚未测试。) 
- 具有邮箱的个人 Microsoft 帐户Outlook.com Microsoft 工作或学校帐户。

如果你没有 Microsoft 帐户，则有两种获取免费帐户的选项：

- 你可以 [注册新的个人 Microsoft 帐户](https://signup.live.com/signup?wa=wsignin1.0&rpsnv=12&ct=1454618383&rver=6.4.6456.0&wp=MBI_SSL_SHARED&wreply=https://mail.live.com/default.aspx&id=64855&cbcxt=mai&bk=1454618383&uiflavor=web&uaid=b213a65b4fdc484382b6622b3ecaa547&mkt=E-US&lc=1033&lic=1)。
- 你可以 [注册 Office 365 开发人员计划](https://developer.microsoft.com/office/dev-program) ，获取免费的 Office 365 订阅。

## <a name="register-a-web-application-with-the-azure-active-directory-admin-center"></a>向 Azure Active Directory 管理中心注册 Web 应用程序

1. 打开浏览器，并转到 [Azure Active Directory 管理中心](https://aad.portal.azure.com)。 使用 **个人帐户**（亦称为“Microsoft 帐户”）或 **工作或学校帐户** 登录。

1. 选择左侧导航栏中的“**Azure Active Directory**”，再选择“**管理**”下的“**应用注册**”。

    ![应用注册的屏幕截图 ](/tutorial/images/app-registrations.png)

1. 选择“新注册”。 在“注册应用”页上，按如下方式设置值。

    - 将“名称”设置为“`Office Add-in Graph Tutorial`”。
    - 将“受支持的帐户类型”设置为“任何组织目录中的帐户和个人 Microsoft 帐户”。
    - 在“重定向 URI”下，将第一个下拉列表设置为“`Single-page application (SPA)`”，并将值设置为“`https://localhost:3000/consent.html`”。

    !["注册应用程序"页的屏幕截图](/tutorial/images/register-an-app.png)

1. 选择“**注册**”。 在 **"Office 外接程序图形** 教程"页上，复制应用程序应用程序 (客户端) **ID** 并保存它，下一步将需要它。

    ![新应用注册的应用程序 ID 屏幕截图](/tutorial/images/application-id.png)

1. 选择“管理”下的“证书和密码”。 选择“新客户端密码”按钮。 在 Description 中 **输入** 值，然后选择"过期"选项之 **一，然后选择**"**添加"。**

1. 离开此页前，先复制客户端密码值。 将在下一步中用到它。

    > [!IMPORTANT]
    > 此客户端密码不会再次显示，所以请务必现在就复制它。

1. 在 **"管理"下****选择** API 权限，然后选择 **"添加权限"。**

1. 选择 **Microsoft Graph，** 然后选择 **委派权限**。

1. 选择以下权限，然后选择 **"添加权限"。**

    - **Calendars.ReadWrite** - 这将允许应用读取和写入用户的日历。
    - **MailboxSettings.Read** - 这将允许应用从用户的邮箱设置获取用户的时区。

    ![已配置权限的屏幕截图](/tutorial/images/configured-permissions.png)

## <a name="configure-office-add-in-single-sign-on"></a>配置 Office 外接程序单一登录

更新应用注册以支持 Office 外接程序[单一登录 (SSO) 。 ](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins)

1. 选择 **"公开 API"。** 在此 **API 定义的** 范围部分中，选择 **"添加范围"。** 当系统提示设置应用程序 **ID URI** 时，将值设置为 `api://localhost:3000/YOUR_APP_ID_HERE` ， `YOUR_APP_ID_HERE` 以应用程序 ID 替换。 选择 **"保存"并继续**。

1. 按如下所示填写字段，然后选择"**添加范围"。**

    - **范围名称：**`access_as_user`
    - **谁可以同意？：管理员和用户**
    - **管理员同意显示名称：**`Access the app as the user`
    - **管理员同意说明：**`Allows Office Add-ins to call the app's web APIs as the current user.`
    - **用户同意显示名称：**`Access the app as you`
    - **用户同意说明：**`Allows Office Add-ins to call the app's web APIs as you.`
    - **状态：已启用**

    !["添加范围"表单的屏幕截图](/tutorial/images/add-scope.png)

1. 在"**授权客户端应用程序**"部分，选择 **"添加客户端应用程序"。** 从以下列表中输入客户端 ID，在"授权范围"下启用范围，然后选择"**添加应用程序"。** 对列表中的每个客户端 ID 重复此过程。

    - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    - `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)
    - `57fb890c-0dab-4253-a5e0-7188c88b2bb4`（Office 网页版）
    - `08e18876-6177-487e-b8b5-cf950c1e598c`（Office 网页版）

## <a name="install-development-certificates"></a>安装开发证书

1. 运行以下命令，为加载项生成和安装开发证书。

    ```Shell
    npx office-addin-dev-certs install
    ```

    如果系统提示确认，请确认操作。 命令完成后，你将看到如下所示的输出。

    ```Shell
    You now have trusted access to https://localhost.
    Certificate: <path>\localhost.crt
    Key: <path>\localhost.key
    ```

1. 将路径复制到 localhost.crt 和 localhost.key，下一步将需要它们。

## <a name="update-the-manifest"></a>更新清单

1. 打开 **manifest.xml** 文件，然后进行以下更改。
    1. 替换为 `NEW_GUID_HERE` 新的 GUID，如 `b4fa03b8-1eb6-4e8b-a380-e0476be9e019` 。
    1. 将应用注册的所有实例 `YOUR_APP_ID_HERE` 替换为应用程序 ID。

## <a name="configure-the-sample"></a>配置示例

1. 将文件 `example.env` 重命名为 `.env` 。
1. 编辑 `.env` 文件并做出以下更改。
    1. 替换为 `YOUR_APP_ID_HERE` 从 **应用** 注册门户获得的应用程序 ID。
    1. 替换为 `YOUR_CLIENT_SECRET_HERE` 从应用注册门户获得的客户密码。
    1. 替换为 `PATH_TO_LOCALHOST.CRT` 命令输出中 localhost.crt 文件 `npx office-addin-dev-certs install` 的路径。
    1. 替换为 `PATH_TO_LOCALHOST.KEY` 命令输出中 localhost.key 文件 `npx office-addin-dev-certs install` 的路径。

1. 将文件 `config.example.js` 重命名为 `config.js` 。
1. 编辑 `config.js` 文件并做出以下更改。
    1. 替换为 `YOUR_APP_ID_HERE` 从 **应用** 注册门户获得的应用程序 ID。
1. 在 CLI (命令行界面) ，导航到此目录并运行以下命令来安装要求。

    ```Shell
    yarn install
    ```

## <a name="run-the-sample"></a>运行示例

1. 在 CLI 中运行以下命令以启动应用程序。

    ```Shell
    yarn start
    ```

1. 在浏览器中 [，转到Office.com](https://www.office.com/) 并登录。 选择 **左侧** 工具栏中的"创建"，然后选择 **"电子表格"。**

    !["创建"菜单的屏幕截图Office.com](/tutorial/images/office-select-excel.png)

1. 选择 **"插入**"选项卡，然后选择 **"Office 加载项"。**

1. 选择 **"上载我的外接程序"，** 然后选择"**浏览"。** 上传manifest.xml **文件** 。

1. 选择 **"主页"****选项卡上的"** 导入日历"按钮以打开任务窗格。

    !["开始"选项卡上"导入日历"按钮的屏幕截图](/tutorial/images/get-started.png)
