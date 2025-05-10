import asyncio
import pyttsx3
from Speech_to_text import record_and_transcribe
from agent import get_agent_reply, ensure_session  # import both functions

engine = pyttsx3.init()

def speak(text):
    print("🤖 Omnitech:", text)
    try:
        engine.say(text)
        engine.runAndWait()
    except KeyboardInterrupt:
        print("🛑 Speech interrupted.")
    except Exception as e:
        print(f"⚠️ TTS error: {e}")

async def main():
    user_id = "user1"
    session_id = "session1"

    # ✅ Ensure session is created at the start
    ensure_session(user_id, session_id)

    try:
        while True:
            print("\n🟢 Say something to ask the assistant (or 'stop' to exit)...")
            spoken_text = record_and_transcribe()

            if spoken_text in ["NO_INPUT", "UNRECOGNIZED", "API_ERROR"]:
                print(f"⚠️ Issue: {spoken_text}.")
                speak("I didn't catch that. Try again or say stop to exit.")
                continue

            print(f"🧠 You said: {spoken_text}")

            if spoken_text.lower() in ["stop", "exit", "shutdown", "quit"]:
                speak("Goodbye!")
                break

            reply = await get_agent_reply(spoken_text, user_id, session_id)
            speak(reply)

    except KeyboardInterrupt:
        print("\n🛑 Interrupted by user.")
        speak("Goodbye!")

if __name__ == "__main__":
    asyncio.run(main())
