import os

# Ollama runs locally - no API key required
DEFAULT_API_KEY = ""

# Available local Ollama models (must be pulled via `ollama pull <model>`)
MODELS = [
    "llama3.2",
    "llama3.1",
    "qwen2.5",
    "mistral",
    "gemma3",
]

DEFAULT_MODEL = "llama3.2"

DEFAULT_SYSTEM_PROMPT = (
    "You are a helpful, precise, and professional conversational AI assistant.\n\n"
    "The user has uploaded one or more documents which are provided as context below. "
    "Analyze the provided document text carefully to answer questions about rules, "
    "specifications, design inputs, etc.\n\n"
    "Guidelines:\n"
    "- Base your answers on the contents of the uploaded documents whenever possible.\n"
    "- If the answer cannot be found in the documents, use your general knowledge, "
    "but clearly state that the information was not in the documents.\n"
    "- Structure your answers using clean Markdown (headings, lists, bold text, "
    "code blocks, or tables) to make them readable.\n"
    "- If the user asks you to extract information, try to present it in a clear "
    "Markdown table format if appropriate."
)

DEFAULT_TEMPERATURE = 0.3
MAX_TOKENS = 4096

KNOWLEDGE_HUB_SYSTEM_PROMPT = (
    "You are a helpful, precise, and professional AI assistant. "
    "Answer questions clearly and accurately using your broad knowledge. "
    "Structure your answers using clean Markdown (headings, lists, bold text, "
    "code blocks, or tables) to make them readable."
)

# Vision model used for PDF OCR (handwriting, scanned pages, embedded images).
# Must be a vision-capable model pulled via `ollama pull <model>`.
# Recommended options: "llava", "llava-phi3", "moondream", "minicpm-v"
VISION_MODEL = "llava"
