import logging
import chainlit as cl
from dotenv import load_dotenv

# Set the logging level to DEBUG
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(message)s',
    datefmt='%d:%m:%Y %H:%M:%S'
)

# Load environment variables from .env file
load_dotenv()

# --- LangChain imports ---
from langchain.memory import ConversationBufferWindowMemory

# --- Custom imports from your environment ---
from langchain_openai import ChatOpenAI
from ExcelTamer.ExcelTamerAgent.AgentBuilder import create_agent

# Create the language model
llm = ChatOpenAI(model_name="gpt-4o-mini", temperature=0)

# Create memory to store the last 10 messages
memory = ConversationBufferWindowMemory(k=10, memory_key="chat_history",
                                        return_messages=True, output_key="output")

# Do not open any specific Excel File; work with currently open Excel Workbook
excel_path = None

# Create agent with memory (assuming your create_agent supports this parameter)

agent = create_agent(excel_path, llm, memory=memory )


@cl.on_chat_start
async def start():
    # Store the agent in user session
    cl.user_session.set("agent", agent)
    await cl.Message(content="Hello! I am your AI assistant. How can I help you today?").send()


@cl.on_message
async def handle_message(message):
    callback_handler = cl.LangchainCallbackHandler()

    # Retrieve the agent from session
    agent_inst = cl.user_session.get("agent")

    # Run the agent asynchronously
    result = await agent_inst.ainvoke({"input": message.content},callbacks=[callback_handler])

    # Extract and display intermediate steps
    intermediate_steps = result.get("intermediate_steps", [])
    for action, observation in intermediate_steps:
        # Display the action taken by the agent
        await cl.Message(content=f"Action: {action.tool}\nInput: {action.tool_input}").send()
        # Display the observation received from the tool
        await cl.Message(content=f"Observation: {observation}").send()

   # Send back the agent's response
    await cl.Message(content=result["output"]).send()


if __name__ == "__main__":
    from chainlit.cli import run_chainlit

    run_chainlit(__file__)
