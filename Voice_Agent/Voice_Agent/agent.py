# agent.py

from google.adk.sessions import Session, InMemorySessionService
from google.adk.agents import LlmAgent
from google.adk.runners import Runner
from google.genai import types

# ✅ Subclass with .get() support
from google.adk.sessions import Session, InMemorySessionService
from dotenv import load_dotenv
import os

load_dotenv()

class PatchedSessionService(InMemorySessionService):
    def get_session(self, user_id: str, session_id: str, app_name: str = None):
        return self.sessions.get(user_id, {}).get(session_id)

    def register(self, session: Session):
        user_id = session.user_id
        session_id = session.id

        if user_id not in self.sessions:
            self.sessions[user_id] = {}

        self.sessions[user_id][session_id] = session




# ✅ Use the patched version
session_service = PatchedSessionService()

# ✅ Shared agent + runner using the SAME session_service
agent = LlmAgent(
    model="gemini-2.0-flash",
    name="voice_agent",
    description="Answers voice questions from the user.",
    instruction="You are a helpful, spoken assistant for voice queries.",
)

runner = Runner(agent=agent, session_service=session_service, app_name="Omnitech_Agent")


def ensure_session(user_id: str, session_id: str):
    if session_service.get_session(user_id, session_id) is None:
        session = Session(
            id=session_id,
            user_id=user_id,
            app_name="Omnitech_Agent"
        )
        session_service.register(session)
        print("✅ Session registered.")
    else:
        print("ℹ️ Session already exists.")





async def get_agent_reply(user_input, user_id="user1", session_id="session1"):
    content = types.Content(role="user", parts=[types.Part(text=user_input)])

    print(f"📦 Sessions before call: {session_service.sessions}")

    async for event in runner.run_async(
        user_id=user_id,
        session_id=session_id,
        new_message=content
    ):
        if event.is_final_response():
            return event.content.parts[0].text
