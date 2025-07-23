from openai import OpenAI

client = OpenAI(
    base_url="https://bio-qwen3-api.brbiotech.tech/v1",
    api_key="no-key-required"
)

response = client.chat.completions.create(
    model="Qwen3-32B",
    # model="/mnt/ai/models/Qwen/Qwen3-32B",
    messages=[{"role": "user", "content": "你好！你是谁？"}],
    max_tokens=512
)
print(response.choices[0].message.content)


