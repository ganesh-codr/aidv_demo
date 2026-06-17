import ollama
from PySide6.QtCore import QThread, Signal


class GroqChatThread(QThread):
    """
    QThread to execute an Ollama chat completion request asynchronously.
    Emits signals for token reception, complete response, and errors.
    """
    token_received = Signal(str)
    finished = Signal(str)
    error = Signal(str)

    def __init__(self, api_key: str, model: str, messages: list, temperature: float = 0.3):
        super().__init__()
        # api_key is unused with Ollama but kept for interface compatibility
        self.model = model
        self.messages = messages
        self.temperature = temperature
        self.is_running = True

    def run(self):
        try:
            stream = ollama.chat(
                model=self.model,
                messages=self.messages,
                stream=True,
                options={"temperature": self.temperature}
            )

            full_response = []
            for chunk in stream:
                if not self.is_running:
                    break
                content = chunk.get("message", {}).get("content", "")
                if content:
                    full_response.append(content)
                    self.token_received.emit(content)

            if self.is_running:
                self.finished.emit("".join(full_response))

        except ollama.ResponseError as e:
            self.error.emit(f"Ollama error: {e.error}")
        except Exception as e:
            self.error.emit(f"Unexpected error: {str(e)}")

    def stop(self):
        """Request the worker thread to stop reading chunks."""
        self.is_running = False
