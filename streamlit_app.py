from __future__ import annotations

import base64
import datetime as dt
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

import streamlit as st

try:
    from openai import OpenAI
except Exception:
    OpenAI = None

try:
    from google import genai
    from google.genai import types as genai_types
except Exception:
    genai = None
    genai_types = None

try:
    from PyPDF2 import PdfReader
except Exception:
    PdfReader = None

try:
    import docx
except Exception:
    docx = None

try:
    import openpyxl
except Exception:
    openpyxl = None

try:
    from pptx import Presentation
except Exception:
    Presentation = None


APP_NAME = "ChatMD"
DEFAULT_PROVIDER = "Google Gemini"

FALLBACK_MODELS: Dict[str, List[str]] = {
    "Google Gemini": [
        "gemini-3-pro-preview",
        "gemini-2.5-pro",
        "gemini-2.5-flash",
        "gemini-2.0-flash",
    ],
    "OpenAI": [
        "gpt-4o",
        "gpt-4o-mini",
        "gpt-4.1",
        "gpt-4.1-mini",
    ],
}

DEFAULT_AGENTS = {
    "Bendras asistentas": {
        "icon": "💬",
        "system_prompt": "Tu esi naudingas, tikslus ir aiškiai lietuviškai atsakantis asistentas.",
        "examples": "",
    },
    "Darbo pagalbininkas": {
        "icon": "🧩",
        "system_prompt": "Padėk struktūruoti užduotis, rašyti aiškiai, trumpai ir profesionaliai lietuviškai.",
        "examples": "Atsakymo struktūra: trumpa išvada, tada konkretūs veiksmai, tada rizikos arba pastabos.",
    },
}


def get_secret(name: str) -> str:
    try:
        return str(st.secrets.get(name, "") or "").strip()
    except Exception:
        return ""


def get_api_key(provider: str) -> str:
    if provider == "Google Gemini":
        return get_secret("GOOGLE_API_KEY")
    if provider == "OpenAI":
        return get_secret("OPENAI_API_KEY")
    return ""


def init_state() -> None:
    st.session_state.setdefault("provider", DEFAULT_PROVIDER)
    st.session_state.setdefault("models_cache", {})
    st.session_state.setdefault("model", FALLBACK_MODELS[DEFAULT_PROVIDER][0])
    st.session_state.setdefault("web_search_enabled", True)
    st.session_state.setdefault("temperature", 0.7)
    st.session_state.setdefault("max_output_tokens", 3000)
    st.session_state.setdefault("mode", "chat")
    st.session_state.setdefault("agent_name", "Bendras asistentas")
    st.session_state.setdefault("agents", DEFAULT_AGENTS.copy())
    st.session_state.setdefault("messages", [])
    st.session_state.setdefault("chat_counter", 1)
    st.session_state.setdefault("history", [])


def reset_chat() -> None:
    if st.session_state.messages:
        title = next((m["content"] for m in st.session_state.messages if m["role"] == "user"), "Pokalbis")
        st.session_state.history.insert(0, {
            "title": title[:60],
            "messages": list(st.session_state.messages),
            "created_at": dt.datetime.now().isoformat(timespec="minutes"),
            "mode": st.session_state.mode,
            "agent": st.session_state.agent_name,
        })
    st.session_state.messages = []
    st.session_state.chat_counter += 1


def runtime_context() -> str:
    today = dt.datetime.now().strftime("%Y-%m-%d")
    return (
        f"Šiandienos data: {today}. "
        "Jei klausimas apie dabartinius įvykius ar faktus, naudok interneto paiešką, kai ji įjungta. "
        "Jei paieška nepavyksta, aiškiai tai pasakyk."
    )


def list_models(provider: str) -> Tuple[List[str], str]:
    api_key = get_api_key(provider)
    if not api_key:
        return FALLBACK_MODELS.get(provider, []), "fallback: nėra API rakto"

    if provider == "Google Gemini":
        if genai is None:
            return FALLBACK_MODELS[provider], "fallback: neįdiegta google-genai"
        try:
            client = genai.Client(api_key=api_key)
            names = []
            for model in client.models.list():
                name = getattr(model, "name", "") or ""
                if name.startswith("models/"):
                    name = name.split("/", 1)[1]
                if name and "embedding" not in name.lower():
                    names.append(name)
            return sorted(set(names)) or FALLBACK_MODELS[provider], "live"
        except Exception as exc:
            return FALLBACK_MODELS[provider], f"fallback: {exc}"

    if provider == "OpenAI":
        if OpenAI is None:
            return FALLBACK_MODELS[provider], "fallback: neįdiegta openai"
        try:
            client = OpenAI(api_key=api_key)
            models = client.models.list()
            names = sorted({m.id for m in models.data if getattr(m, "id", None)})
            names = [m for m in names if m.startswith(("gpt-", "o", "chatgpt-"))]
            return names or FALLBACK_MODELS[provider], "live"
        except Exception as exc:
            return FALLBACK_MODELS[provider], f"fallback: {exc}"

    return [], "fallback"


def read_uploaded_file(uploaded_file) -> str:
    name = uploaded_file.name
    suffix = PathLike(name).suffix
    data = uploaded_file.getvalue()

    try:
        if suffix in {".png", ".jpg", ".jpeg", ".webp", ".gif"}:
            encoded = base64.b64encode(data).decode("utf-8")
            return f"[Vaizdas {name}, base64 pradžia]: {encoded[:3000]}"

        if suffix == ".pdf":
            if PdfReader is None:
                return f"[{name}] PDF nuskaitymui reikia PyPDF2."
            import io
            reader = PdfReader(io.BytesIO(data))
            pages = [(p.extract_text() or "") for p in reader.pages[:20]]
            return f"[{name}]\n" + "\n".join(pages)[:15000]

        if suffix == ".docx":
            if docx is None:
                return f"[{name}] DOCX nuskaitymui reikia python-docx."
            import io
            document = docx.Document(io.BytesIO(data))
            text = "\n".join(p.text for p in document.paragraphs if p.text.strip())
            return f"[{name}]\n{text[:15000]}"

        if suffix == ".xlsx":
            if openpyxl is None:
                return f"[{name}] XLSX nuskaitymui reikia openpyxl."
            import io
            wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
            rows = []
            for ws in wb.worksheets[:3]:
                rows.append(f"Lapas: {ws.title}")
                for row in ws.iter_rows(max_row=80, values_only=True):
                    vals = [str(v) for v in row if v is not None]
                    if vals:
                        rows.append(" | ".join(vals))
            return f"[{name}]\n" + "\n".join(rows)[:15000]

        if suffix == ".pptx":
            if Presentation is None:
                return f"[{name}] PPTX nuskaitymui reikia python-pptx."
            import io
            prs = Presentation(io.BytesIO(data))
            slides = []
            for i, slide in enumerate(prs.slides[:30], 1):
                texts = []
                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text:
                        texts.append(shape.text)
                if texts:
                    slides.append(f"Skaidrė {i}:\n" + "\n".join(texts))
            return f"[{name}]\n" + "\n\n".join(slides)[:15000]

        return f"[{name}]\n" + data.decode("utf-8", errors="ignore")[:15000]
    except Exception as exc:
        return f"[{name}] Nepavyko nuskaityti: {exc}"


class PathLike:
    def __init__(self, name: str) -> None:
        self.name = name

    @property
    def suffix(self) -> str:
        if "." not in self.name:
            return ""
        return "." + self.name.rsplit(".", 1)[1].lower()


def build_system_prompt() -> str:
    base = runtime_context()
    if st.session_state.mode == "agent":
        agent = st.session_state.agents.get(st.session_state.agent_name, DEFAULT_AGENTS["Bendras asistentas"])
        prompt = agent.get("system_prompt", "")
        examples = agent.get("examples", "")
        if examples.strip():
            prompt += "\n\nAgento pavyzdžiai / žinios:\n" + examples.strip()
        return base + "\n\n" + prompt
    return base + "\n\nTu esi naudingas, tikslus ir aiškiai lietuviškai atsakantis asistentas."


def answer_google(model: str, messages: List[Dict[str, str]], files_context: str) -> str:
    api_key = get_api_key("Google Gemini")
    if not api_key:
        return "🔴 Google Gemini API raktas neįvestas Streamlit Secrets."

    client = genai.Client(api_key=api_key)
    parts = [f"Sisteminė instrukcija:\n{build_system_prompt()}"]
    for m in messages[-20:]:
        role = "Vartotojas" if m["role"] == "user" else "Asistentas"
        parts.append(f"{role}: {m['content']}")
    if files_context:
        parts.append("Prisegtų failų turinys:\n" + files_context)

    config = {
        "temperature": st.session_state.temperature,
        "max_output_tokens": int(st.session_state.max_output_tokens),
    }

    if st.session_state.web_search_enabled and genai_types is not None:
        try:
            config = genai_types.GenerateContentConfig(
                temperature=st.session_state.temperature,
                max_output_tokens=int(st.session_state.max_output_tokens),
                tools=[genai_types.Tool(google_search=genai_types.GoogleSearch())],
            )
        except Exception:
            pass

    response = client.models.generate_content(
        model=model,
        contents="\n\n".join(parts),
        config=config,
    )
    return getattr(response, "text", "") or "Gautas tuščias atsakymas."


def answer_openai(model: str, messages: List[Dict[str, str]], files_context: str) -> str:
    api_key = get_api_key("OpenAI")
    if not api_key:
        return "🔴 OpenAI API raktas neįvestas Streamlit Secrets."

    client = OpenAI(api_key=api_key)
    input_messages = [{"role": "system", "content": build_system_prompt()}]
    input_messages.extend(messages[-20:])
    if files_context:
        input_messages.append({"role": "user", "content": "Prisegtų failų turinys:\n" + files_context})

    # Kai kurie nauji OpenAI reasoning modeliai nepalaiko temperature parametro.
    # Todėl pirmiausia siunčiame saugų minimalų užklausos variantą be temperature.
    kwargs = {
        "model": model,
        "input": input_messages,
        "max_output_tokens": int(st.session_state.max_output_tokens),
    }

    # Temperature dedame tik modeliams, kurie dažniausiai jį palaiko.
    if not model.startswith(("gpt-5", "o1", "o3", "o4")):
        kwargs["temperature"] = st.session_state.temperature

    if st.session_state.web_search_enabled:
        kwargs["tools"] = [{"type": "web_search"}]

    try:
        response = client.responses.create(**kwargs)
    except Exception as exc:
        msg = str(exc)
        # Jei modelis vis tiek atmetė temperature, kartojame be jo.
        if "temperature" in msg and "Unsupported parameter" in msg:
            kwargs.pop("temperature", None)
            response = client.responses.create(**kwargs)
        # Jei web_search tipas konkrečiam SDK / modeliui netinka, kartojame be paieškos,
        # kad pats pokalbis nesulūžtų.
        elif "web_search" in msg or "tool" in msg.lower():
            kwargs.pop("tools", None)
            response = client.responses.create(**kwargs)
        else:
            raise

    return getattr(response, "output_text", "") or "Gautas tuščias atsakymas."


def get_answer(prompt: str, uploaded_files) -> str:
    files_context = ""
    if uploaded_files:
        chunks = [read_uploaded_file(f) for f in uploaded_files]
        files_context = "\n\n".join(chunks)

    provider = st.session_state.provider
    model = st.session_state.model

    try:
        if provider == "Google Gemini":
            return answer_google(model, st.session_state.messages, files_context)
        if provider == "OpenAI":
            return answer_openai(model, st.session_state.messages, files_context)
        return "Nežinomas teikėjas."
    except Exception as exc:
        return f"⚠️ Klaida: {exc}"


def render_css() -> None:
    st.markdown(
        """
        <style>
        .stApp {
            background: #f7f9fc;
            color: #111827;
        }
        [data-testid="stSidebar"] {
            background: #eef3fa;
            border-right: 1px solid #d8e2f0;
        }
        .chat-title {
            display: flex;
            align-items: center;
            gap: 0.8rem;
            padding: 0.4rem 0 1rem 0;
        }
        .chat-logo {
            width: 58px;
            height: 58px;
            object-fit: contain;
        }
        .status-pill {
            display:inline-block;
            padding: 0.2rem 0.55rem;
            border-radius: 999px;
            background:#dbeafe;
            color:#1d4ed8;
            font-size:0.82rem;
            margin-right:0.35rem;
            margin-bottom:0.25rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_brand() -> None:
    logo_path = "chatmd_logo.png"
    logo_html = ""
    try:
        with open(logo_path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode("utf-8")
        logo_html = f'<img class="chat-logo" src="data:image/png;base64,{b64}" />'
    except Exception:
        logo_html = ""
    st.sidebar.markdown(
        f"""
        <div class="chat-title">
            {logo_html}
            <div>
                <div style="font-size:1.75rem;font-weight:700;line-height:1;">Chat</div>
                <div style="font-size:0.85rem;color:#526173;">AI pokalbiai ir agentai</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_sidebar() -> None:
    render_brand()

    if st.sidebar.button("＋ Naujas pokalbis", use_container_width=True):
        reset_chat()
        st.rerun()

    st.sidebar.divider()

    st.sidebar.subheader("Režimas")
    mode = st.sidebar.radio(
        "Pasirink režimą",
        ["chat", "agent"],
        format_func=lambda x: "💬 Chat" if x == "chat" else "🤖 Agentas",
        horizontal=True,
        label_visibility="collapsed",
    )
    if mode != st.session_state.mode:
        reset_chat()
        st.session_state.mode = mode
        st.rerun()

    if st.session_state.mode == "agent":
        names = list(st.session_state.agents.keys())
        selected = st.sidebar.selectbox("Aktyvus agentas", names, index=names.index(st.session_state.agent_name))
        if selected != st.session_state.agent_name:
            reset_chat()
            st.session_state.agent_name = selected
            st.rerun()

        with st.sidebar.expander("Kurti / redaguoti agentą"):
            current = st.session_state.agents[st.session_state.agent_name]
            name = st.text_input("Pavadinimas", st.session_state.agent_name)
            icon = st.text_input("Ikona", current.get("icon", "🤖"))
            system_prompt = st.text_area("System Prompt", current.get("system_prompt", ""), height=120)
            examples = st.text_area("Pavyzdžiai / žinios", current.get("examples", ""), height=160)
            if st.button("Išsaugoti agentą"):
                st.session_state.agents[name] = {
                    "icon": icon,
                    "system_prompt": system_prompt,
                    "examples": examples,
                }
                st.session_state.agent_name = name
                reset_chat()
                st.rerun()

    st.sidebar.divider()

    st.sidebar.subheader("AI nustatymai")
    provider = st.sidebar.selectbox("Teikėjas", ["Google Gemini", "OpenAI"], index=["Google Gemini", "OpenAI"].index(st.session_state.provider))
    if provider != st.session_state.provider:
        st.session_state.provider = provider
        models, _ = list_models(provider)
        st.session_state.model = models[0] if models else ""
        reset_chat()
        st.rerun()

    key = get_api_key(st.session_state.provider)
    st.sidebar.markdown("🟢 API raktas aktyvus" if key else "🔴 API raktas neįvestas")

    if st.sidebar.button("Atnaujinti modelių sąrašą", use_container_width=True):
        models, source = list_models(st.session_state.provider)
        st.session_state.models_cache[st.session_state.provider] = models
        if st.session_state.model not in models:
            st.session_state.model = models[0] if models else ""
        st.sidebar.info(f"Modeliai: {source}")

    models = st.session_state.models_cache.get(st.session_state.provider)
    if not models:
        models, _ = list_models(st.session_state.provider)
        st.session_state.models_cache[st.session_state.provider] = models

    if st.session_state.model not in models and models:
        st.session_state.model = models[0]

    st.session_state.model = st.sidebar.selectbox("Modelis", models, index=models.index(st.session_state.model) if st.session_state.model in models else 0)
    st.session_state.web_search_enabled = st.sidebar.toggle("Interneto paieška", value=st.session_state.web_search_enabled)
    st.session_state.temperature = st.sidebar.slider("Kūrybiškumas", 0.0, 2.0, float(st.session_state.temperature), 0.1)
    st.session_state.max_output_tokens = st.sidebar.slider("Max output tokens", 512, 8000, int(st.session_state.max_output_tokens), 256)

    st.sidebar.divider()
    st.sidebar.subheader("Istorija šioje sesijoje")
    if not st.session_state.history:
        st.sidebar.caption("Istorija atsiras pradėjus naują pokalbį.")
    for i, chat in enumerate(st.session_state.history[:15]):
        col1, col2 = st.sidebar.columns([0.82, 0.18])
        with col1:
            if st.button(chat["title"], key=f"hist_{i}", use_container_width=True):
                st.session_state.messages = list(chat["messages"])
                st.session_state.mode = chat.get("mode", "chat")
                st.session_state.agent_name = chat.get("agent", "Bendras asistentas")
                st.rerun()
        with col2:
            if st.button("🗑", key=f"del_{i}"):
                st.session_state.history.pop(i)
                st.rerun()


def main() -> None:
    st.set_page_config(
        page_title=APP_NAME,
        page_icon="💬",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    init_state()
    render_css()
    render_sidebar()

    st.title("ChatMD")
    st.markdown(
        f"""
        <span class="status-pill">{st.session_state.provider}</span>
        <span class="status-pill">{st.session_state.model}</span>
        <span class="status-pill">{"Agentas: " + st.session_state.agent_name if st.session_state.mode == "agent" else "Paprastas chat"}</span>
        <span class="status-pill">Internetas: {"ON" if st.session_state.web_search_enabled else "OFF"}</span>
        """,
        unsafe_allow_html=True,
    )

    uploaded_files = st.file_uploader(
        "Prisegti failus",
        type=["png", "jpg", "jpeg", "webp", "gif", "pdf", "docx", "xlsx", "pptx", "txt", "csv", "md", "py", "json"],
        accept_multiple_files=True,
    )

    for m in st.session_state.messages:
        with st.chat_message(m["role"]):
            st.markdown(m["content"])

    prompt = st.chat_input("Rašykite žinutę lietuviškai...")
    if prompt:
        st.session_state.messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)

        with st.chat_message("assistant"):
            with st.spinner("Agentas rašo..."):
                answer = get_answer(prompt, uploaded_files)
                st.markdown(answer)

        st.session_state.messages.append({"role": "assistant", "content": answer})


if __name__ == "__main__":
    main()
