# deli_api

deli_api 是一个用于 DeliLegal 开放 API 的轻量级集成包。它封装了两个上游接口：

- `queryListLaw`：法律法规检索

- `queryListCase`：案例检索

该包提供三个层级的功能：

- `DeliLegalClient`：具有标准化输出的底层 HTTP 客户端

- `LawRetriever` 和 `CaseRetriever`：LangChain 检索器

- `create_search_laws_tool` 和 `create_search_cases_tool`：LangChain 工具

# 1. 安装

激活 conda 环境，并以可编辑模式安装该包：

```powershell
conda activate YiZhi
pip install -e .
```

# 2. 环境变量

从模板创建本地 .env 文件：

```powershell
Copy-Item .env.example .env
```

必填变量：

- `DELILEGAL_APP_ID`

- `DELILEGAL_SECRET`

可选变量：

- `DELILEGAL_BASE_URL`
        

默认值：`https://openapi.delilegal.com`

- `DELILEGAL_TIMEOUT`
        

默认值：`30`

# 3. 客户端使用

客户端是最底层的入口，主要处理以下内容：

- 认证请求头

- HTTP 请求发送

- 上游业务异常

- JSON 解析

- 记录提取

- 结果标准化

示例：

```python
from deli_api import init_case_client

with init_case_client(".env") as client:
    case_hits = client.search_cases(
        query="上班途中车祸工伤案例",
        page_size=3,
    )
```

如需类方法写法，仍可使用：

```python
from deli_api import DeliLegalClient

with DeliLegalClient.from_env(".env") as client:
    law_hits = client.search_laws(
        query="深圳市房地产相关的法律规定有哪些？",
        page_size=3,
        time_liness_type_arr=["5"],
        field_name="semantic",
    )

    case_hits = client.search_cases(
        query="上班途中车祸工伤案例",
        page_size=3,
    )

    for hit in law_hits:
        print(hit.title, hit.citation)

    for hit in case_hits:
        print(hit.title, hit.citation)
```

# 4. LangChain 检索器使用

当需要将上游结果作为 `Document` 对象用于 RAG（检索增强生成）时，可使用检索器。

```python
from deli_api import CaseRetriever, DeliLegalClient, LawRetriever

with DeliLegalClient.from_env(".env") as client:
    law_retriever = LawRetriever(
        client=client,
        page_size=3,
        time_liness_type_arr=["5"],
        field_name="semantic",
    )

    case_retriever = CaseRetriever(
        client=client,
        page_size=3,
    )

    law_docs = law_retriever.invoke("深圳市房地产相关的法律规定有哪些？")
    case_docs = case_retriever.invoke("上班途中车祸工伤案例")

    print(law_docs[0].metadata)
    print(case_docs[0].page_content[:200])
```

每个 `Document.metadata` 包含标准化字段，例如：

- `source_type`（来源类型）

- `title`（标题）

- `citation`（引用信息）

- `score`（得分）

- `source_id`（来源ID）

- 案例专属元数据，如 `case_no`（案号）、`court_name`（法院名称）、`cause`（案由）

- 法规专属元数据，如 `law_name`（法规名称）、`publish_date`（发布日期）、`active_date`（生效日期）

# 5. LangChain 工具使用

当需要让智能体决定何时调用 DeliLegal 检索 API 时，可使用工具。

```python
from deli_api import DeliLegalClient, create_search_cases_tool, create_search_laws_tool

with DeliLegalClient.from_env(".env") as client:
    law_tool = create_search_laws_tool(
        client=client,
        defaults={"page_size": 3, "field_name": "semantic"},
    )

    case_tool = create_search_cases_tool(
        client=client,
        defaults={"page_size": 3},
    )

    print(law_tool.invoke({"query": "深圳市房地产相关的法律规定有哪些？"}))
    print(case_tool.invoke({"query": "上班途中车祸工伤案例"}))
```

工具输出为 JSON 字符串，便于回传给智能体。

# 6. 返回数据模型

`search_laws` 和 `search_cases` 方法均返回 `list[LegalHit]`（LegalHit 对象列表）。

LegalHit 的主要字段：

- `source_type`（来源类型）

- `title`（标题）

- `content`（内容）

- `score`（得分）

- `source_id`（来源ID）

- `citation`（引用信息）

- `url`（链接）

- `metadata`（元数据）

- `raw`（原始数据）

若需要 LangChain 的 `Document` 对象，可调用：

```python
document = hit.to_document()
```

# 7. 异常体系

该包提供以下异常层级结构：

- `DeliLegalError`（基础异常）

- `DeliLegalConfigError`（配置异常）

- `DeliLegalRequestError`（请求异常）

- `DeliLegalHTTPStatusError`（HTTP 状态异常）

- `DeliLegalResponseDecodeError`（响应解码异常）

- `DeliLegalResponseFormatError`（响应格式异常）

- `DeliLegalAPIError`（API 异常）

- `DeliLegalAuthenticationError`（认证异常）

- `DeliLegalUpstreamError`（上游服务异常）

示例：

```python
from deli_api import (
    DeliLegalAuthenticationError,
    DeliLegalClient,
    DeliLegalConfigError,
    DeliLegalError,
)

try:
    with DeliLegalClient.from_env(".env") as client:
        hits = client.search_laws(query="深圳市房地产相关的法律规定有哪些？")
except DeliLegalConfigError as exc:
    print("环境配置不完整:", exc.to_dict())
except DeliLegalAuthenticationError as exc:
    print("认证失败:", exc.to_dict())
except DeliLegalError as exc:
    print("其他 Deli 相关异常:", exc.to_dict())
```

每个异常实例都可通过 `to_dict()` 方法获取结构化上下文信息，包含：

- `message`（异常信息）

- `code`（异常码）

- `status_code`（状态码）

- `path`（请求路径）

- `details`（详细信息）

- `payload`（请求负载）

- `response_text`（响应文本）

# 8. 推荐集成模式

对于法律智能体项目，建议采用以下清晰的层级结构：

1. 将 `DeliLegalClient` 作为唯一知晓上游 API 结构的模块

2. 在 RAG 流水线中使用检索器

3. 在智能体工作流中使用工具

4. 将提示词逻辑放在该包外部

5. 将下游结果后处理逻辑放在该包外部

这样可使 `deli_api` 专注于传输、结果标准化和 LangChain 集成功能。

# 9. 本地冒烟测试

配置好 .env 文件后，运行以下命令进行冒烟测试：

```powershell
python examples/basic_usage.py
```

该测试将执行以下操作：

- 调用法规 API 进行检索

- 调用案例 API 进行检索

- 运行两个检索器

- 运行两个工具

# 10. 源文件说明

- `src/deli_api/client.py`（客户端实现）

- `src/deli_api/exceptions.py`（异常定义）

- `src/deli_api/schemas.py`（数据模型定义）

- `src/deli_api/retrievers.py`（检索器实现）

- `src/deli_api/tools.py`（工具实现）

- `examples/basic_usage.py`（示例代码）
