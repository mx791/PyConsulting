import json
import os.path
from collections.abc import Callable
import boto3


MODEL_ID = "anthropic.claude-3-sonnet-20240229-v1:0"


def load_prompt(prompt: str) -> str:
    """Execute a prompt on AWS bedrock and returns LLM's answer."""

    contentType = "application/json"
    accept = "application/json"

    prompt_config = {
        "anthropic_version": "bedrock-2023-05-31",
        "max_tokens": 1000,
        "temperature": 0.0,
        "messages": [
            {
                "role": "user",
                "content": prompt,
            }
        ],
    }
    body = json.dumps(prompt_config)

    bedrock = boto3.client(
        service_name="bedrock-runtime",
        region_name="eu-west-3",
        aws_access_key_id=os.getenv("AWS_ACCESS_KEY"),
        aws_secret_access_key=os.getenv("AWS_SECRET_KEY"),
    )

    response = bedrock.invoke_model(
        modelId=MODEL_ID, contentType=contentType, accept=accept, body=body
    )
    return json.loads(response.get("body").read())["content"][0]["text"]


def cache(key: str, fc: Callable[[], str]) -> str:
    """Creates a local cache to save LLM's answer."""

    fname = f"./cache/{key}.txt"
    if os.path.isfile(fname):
        file = open(fname, "r")
        txt = file.read()
        file.close()
    else:
        print(f"cache key: {key} not found, fetching data from {MODEL_ID}")
        txt = fc()
        file = open(fname, "w+")
        file.write(txt)
        file.close()
    return txt


def create_company_summary(company_name: str) -> str:
    """Ask the LLM to introduce the company."""

    txt = cache(
        company_name,
        lambda: load_prompt(
            f"""Create a report to present the company: {company_name} to a financial investor. The report will be formated like the following:
Introduction:
[a small introduction within 2 or 3 sentences]

Key points:
[5 bullet points]
"""
        ),
    )

    # post formatting

    # remove everything that stands before "introduction"
    if "Introduction:" in txt:
        txt = txt.split("Introduction:")[1]

    txt = txt.replace("\n\n", "\n")
    txt = txt.replace("•", "-")
    txt = txt.replace("Key points:", "")
    return txt


def top_companies_to_invest(company_list: list[str]) -> str:
    """Ask the LLM for recommendations."""

    prompt = f"""As a financial professional, within the following companies, which would be the best 3 to invest in, and why ?
{['- ' + c + '\n' for c in company_list]}
Give your answer like this:
[first company name] : [justification]
[second company name] : [justification]
[third company name] : [justification]
"""
    txt = cache("top_invest", lambda: load_prompt(prompt))
    return txt
