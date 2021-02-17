---
ms.openlocfilehash: facdbb5c42e60e5bb0ee98b06ef68939a6010a67
ms.sourcegitcommit: 24bb4b3df6a035806a58b609e1d8078ac5505fa1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/17/2021
ms.locfileid: "50274214"
---
<!-- markdownlint-disable MD002 MD041 -->

在此部分中，您将添加在用户日历上创建事件的能力。

## <a name="implement-the-api"></a>实现 API

1. 打开 **./src/api/graph.ts** 并添加以下代码，以实施新的事件 API `POST /graph/newevent` () 。

    :::code language="typescript" source="../demo/graph-tutorial/src/api/graph.ts" id="CreateEventSnippet":::

1. 打开 **./src/addin/taskpane.js** 并添加以下函数以调用新的事件 API。

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="CreateEventSnippet":::

1. 保存所有更改，重新启动服务器，并刷新 Excel 中的任务窗格 (关闭任何打开的任务窗格，然后重新打开) 。

    ![创建事件表单的屏幕截图](images/create-event-ui.png)

1. 填写表单，然后选择"创建 **"。** 验证事件是否添加到用户日历。
