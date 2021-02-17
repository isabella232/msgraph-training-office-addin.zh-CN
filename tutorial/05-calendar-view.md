---
ms.openlocfilehash: ade142f1518d9bd56fa1472889721a4c8995bc3c
ms.sourcegitcommit: 24bb4b3df6a035806a58b609e1d8078ac5505fa1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/17/2021
ms.locfileid: "50274215"
---
<!-- markdownlint-disable MD002 MD041 -->

在此练习中，你将将 Microsoft Graph 合并到应用程序中。 对于此应用程序，你将使用 [microsoft-graph-client](https://github.com/microsoftgraph/msgraph-sdk-javascript) 库调用 Microsoft Graph。

## <a name="get-calendar-events-from-outlook"></a>从 Outlook 获取日历事件

首先添加 API， [从用户的](https://docs.microsoft.com/graph/api/user-list-calendarview) 日历获取日历视图。

1. 打开 **./src/api/graph.ts，** 将以下语句 `import` 添加到文件顶部。

    ```typescript
    import { zonedTimeToUtc } from 'date-fns-tz';
    import { findOneIana } from 'windows-iana';
    import * as graph from '@microsoft/microsoft-graph-client';
    import { Event, MailboxSettings } from 'microsoft-graph';
    import 'isomorphic-fetch';
    import { getTokenOnBehalfOf } from './auth';
    ```

1. 添加以下函数以初始化 Microsoft Graph SDK 并返回 **客户端**。

    :::code language="typescript" source="../demo/graph-tutorial/src/api/graph.ts" id="GetClientSnippet":::

1. 添加以下函数，从用户的邮箱设置获取用户的时区，并将该值转换为 IANA 时区标识符。

    :::code language="typescript" source="../demo/graph-tutorial/src/api/graph.ts" id="GetTimeZonesSnippet":::

1. 将以下函数 (`const graphRouter = Router();` 行) 实现 API 终结点 `GET /graph/calendarview` () 。

    :::code language="typescript" source="../demo/graph-tutorial/src/api/graph.ts" id="GetCalendarViewSnippet":::

    考虑此代码执行什么功能。

    - 它获取用户的时区，并使用该时区将请求的日历视图的开始和结尾转换为 UTC 值。
    - 它会对 `GET` Graph `/me/calendarview` API 终结点进行调用。
        - 它使用该函数来设置标头，从而根据用户的时区调整返回事件的开始时间和 `header` `Prefer: outlook.timezone` 结束时间。
        - 它使用该 `query` 函数添加 and `startDateTime` `endDateTime` 参数，并设置日历视图的开始和结束。
        - 它使用该 `select` 函数仅请求外接程序使用的字段。
        - 它使用 `orderby` 函数按开始时间对结果进行排序。
        - 它使用该 `top` 函数将单个请求中的结果限制为 25。
    - 它使用 **PageIteratorCallback** 对象 [](https://docs.microsoft.com/graph/sdks/paging)来对结果进行引用，并可在更多结果页面可用时提出其他请求。

## <a name="update-the-ui"></a>更新 UI

现在，让我们更新任务窗格，以允许用户指定日历视图的开始日期和结束日期。

1. 打开 **./src/addin/taskpane.js，** 将 `showMainUi` 现有函数替换为以下内容。

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="MainUiSnippet":::

    此代码添加一个简单的表单，以便用户可以指定开始日期和结束日期。 它还实现用于创建新事件的第二个表单。 该表单目前未执行任何功能，您将在下一节中实现该功能。

1. 将以下代码添加到文件以在活动工作表中创建一个表，其中包含从日历视图中检索到的事件。

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="WriteToSheetSnippet":::

1. 添加以下函数以调用日历视图 API。

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="GetCalendarSnippet":::

1. 保存所有更改，重新启动服务器，并刷新 Excel 中的任务窗格 (关闭任何打开的任务窗格，然后重新打开) 。

    ![导入表单的屏幕截图](images/get-calendar-view-ui.png)

1. 选择开始日期和结束日期，然后选择"**导入"。**

    ![事件表的屏幕截图](images/calendar-view-table.png)
