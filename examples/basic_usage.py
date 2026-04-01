from deli_api import (
    CaseRetriever,
    DeliLegalClient,
    LawRetriever,
    create_search_cases_tool,
    create_search_laws_tool,
)


def main() -> None:
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

        law_query = "\u6df1\u5733\u5e02\u623f\u5730\u4ea7\u76f8\u5173\u7684\u6cd5\u5f8b\u89c4\u5b9a\u6709\u54ea\u4e9b\uff1f"
        case_query = "\u4e0a\u73ed\u9014\u4e2d\u8f66\u7978\u5de5\u4f24\u6848\u4f8b"

        law_docs = law_retriever.invoke(law_query)
        print(f"Law retriever results: {len(law_docs)}")
        for doc in law_docs:
            print(doc.metadata.get("citation", doc.metadata.get("title")))

        case_docs = case_retriever.invoke(case_query)
        print(f"\nCase retriever results: {len(case_docs)}")
        for doc in case_docs:
            print(doc.metadata.get("citation", doc.metadata.get("title")))

        law_tool = create_search_laws_tool(client=client, defaults={"page_size": 3})
        case_tool = create_search_cases_tool(client=client, defaults={"page_size": 3})

        print("\nLaw tool preview:")
        print(law_tool.invoke({"query": law_query}))

        print("\nCase tool preview:")
        print(case_tool.invoke({"query": case_query}))


if __name__ == "__main__":
    main()
