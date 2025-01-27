import json

from langchain.agents import initialize_agent, AgentType
from langchain_core.tools import tool
from langchain_openai import ChatOpenAI

from dotenv import load_dotenv

load_dotenv()

@tool
def calculate_all(a: float, b: float) -> dict:
    """
    Calculate sum, difference, product, and quotient of two numbers.
    """
    return {
        'sum':a + b,
        'difference':a - b,
        'product':a * b,
        'quotient':(a / b) if b != 0 else float("inf")
    }


llm = ChatOpenAI(temperature=0)
agent = initialize_agent(
    tools=[calculate_all],
    llm=llm,
    agent=AgentType.OPENAI_FUNCTIONS,
    verbose=True
)

response = agent.run("Calculate all operations for 12 and 4.")
print("Agent gave this final response:", response)

# If you want just the JSON payload from the tool, you could do:
# (depends on how the agent is configured)
try:
    data = json.loads(response)
    print("Parsed JSON data:", data)
except json.JSONDecodeError:
    print("Agent did not return pure JSON.")