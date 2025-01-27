import chainlit as cl

from SimpleExcelAgent import agent

@cl.on_chat_start
async def start():
    cl.user_session.set("agent", agent)
    await cl.Message(content="Hello! I am your AI assistant. How can I help you today?").send()

@cl.on_message
async def handle_message(message):
    agent_inst = cl.user_session.get("agent")
    res = await agent_inst.arun(
        message.content, callbacks=[cl.AsyncLangchainCallbackHandler()]
    )
    await cl.Message(content=res).send()

if __name__ == "__main__":
    from chainlit.cli import run_chainlit
    run_chainlit(__file__)