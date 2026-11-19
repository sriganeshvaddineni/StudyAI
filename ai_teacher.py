
import streamlit as st
from google import genai
from ddgs import DDGS
import time


# ==========================================
# GEMINI MODELS
# ==========================================

MODEL_NAMES = [
    "gemini-3.1-flash-lite",
    "gemini-2.5-flash-lite",
    "gemini-3.5-flash-lite"
]


def get_gemini_client():
    return genai.Client(
        api_key=st.secrets["GEMINI_API_KEY"]
    )


# ==========================================
# FREE ONLINE SEARCH
# ==========================================

def search_online(topic, subject, unit):

    query = f"{topic} {subject} {unit}"

    backends = ["bing", "yahoo"]

    for backend in backends:

        try:
            results = DDGS().text(
                query,
                backend=backend,
                max_results=5
            )

            if results:
                return results

        except Exception as error:
            print(f"{backend} search failed: {error}")

    return []


# ==========================================
# AI LESSON GENERATION
# ==========================================

def generate_lesson(
    topic,
    subject,
    unit,
    mode,
    material_text=""
):

    # Search online for every topic
    online_results = search_online(
        topic,
        subject,
        unit
    )
    print("ONLINE SEARCH RESULTS:", len(online_results))
    

    # ======================================
    # PREPARE ONLINE INFORMATION
    # ======================================

    online_text = ""

    if online_results:

        online_text = "\nONLINE SEARCH RESULTS:\n"

        for index, result in enumerate(
            online_results,
            start=1
        ):

            online_text += f"""
SOURCE {index}
Title: {result.get("title", "")}
Description: {result.get("body", "")}
URL: {result.get("href", "")}
"""

    else:

        online_text = """
Online search did not return results.
Do not invent online information or sources.
Use uploaded materials and general knowledge.
"""


    # ======================================
    # PREPARE UPLOADED MATERIAL
    # ======================================

    if material_text and material_text.strip():

        material_instructions = f"""
UPLOADED STUDY MATERIAL:

{material_text}

Instructions:
1. Give priority to the uploaded study material.
2. Use relevant information from the material.
3. Combine it with useful information from online search.
4. Do not invent facts and attribute them to the uploaded material.
5. Clearly identify additional explanations where useful.
6. Explain the content instead of copying large portions.
"""

    else:

        material_instructions = """
No study material was uploaded.

Use online search results and general knowledge.
Do not claim that uploaded materials were provided.
"""


    # ======================================
    # COMMON SOURCE INSTRUCTIONS
    # ======================================

    source_instructions = f"""
SOURCE USAGE RULES:

1. Use the online search results as supporting information.
2. Prefer trustworthy educational and official sources.
3. Do not blindly trust search snippets.
4. Do not invent facts, URLs or source names.
5. If sources disagree, mention the uncertainty.
6. The actual source links will be displayed separately.

{online_text}

{material_instructions}
"""


    # ======================================
    # LEARNING MODES
    # ======================================

    prompts = {

        "telugu": f"""
Teach "{topic}" from subject "{subject}",
unit "{unit}".

LANGUAGE RULES:

1. Use Telugu ONLY in English letters (Roman Telugu).
2. NEVER use Telugu script characters.
3. Explain naturally like friends chatting on WhatsApp.
4. Mix simple English technical terms with Roman Telugu.
5. Use English headings.
6. Keep programming keywords and code in English.
7. Explain difficult concepts simply.
8. Include practical examples.
9. Make the explanation beginner-friendly.
10. Give important points and a conclusion.

Example style:

"Python anedi oka easy programming language.
Manam computer ki instructions ivvadaniki
Python use chestham."

Use Roman Telugu and English ONLY.

{source_instructions}
""",

        "simple": f"""
Explain "{topic}" from "{subject}",
unit "{unit}" in simple English.

Include:
1. Definition
2. Easy explanation
3. Important points
4. Step-by-step explanation
5. Real-world examples
6. Applications
7. Advantages and disadvantages if relevant
8. Conclusion

Assume the student is a complete beginner.

{source_instructions}
""",

        "exam": f"""
Prepare an exam-ready answer for "{topic}"
from "{subject}", unit "{unit}".

Include:
1. Definition
2. Introduction
3. Important points with headings
4. Detailed explanation
5. Suitable examples
6. Advantages and disadvantages if relevant
7. Applications
8. Conclusion

Use academic English and numbered points.
Prepare a detailed 5 to 10 mark answer.

{source_instructions}
""",

        "practice": f"""
Create practice material for "{topic}"
from "{subject}", unit "{unit}".

SECTION A: MCQs
- Five questions.
- Four options each.
- Correct answer.
- Brief explanation.

SECTION B: SHORT ANSWERS
- Three questions with answers.

SECTION C: LONG ANSWERS
- Two questions with detailed answers.

Use simple English.
Focus on the uploaded material and online
information where relevant.

{source_instructions}
"""
    }


    prompt = prompts.get(mode)

    if not prompt:
        return "Invalid learning mode."


    # ======================================
    # GENERATE WITH GEMINI
    # ======================================

    try:

        client = get_gemini_client()
        errors = []

        for model_name in MODEL_NAMES:

            for attempt in range(2):

                try:

                    response = client.models.generate_content(
                        model=model_name,
                        contents=prompt
                    )

                    if response.text:

                        answer = response.text

                        # Append actual search results.
                        if online_results:

                            answer += "\n\n---\n"
                            answer += "### Online Sources\n\n"

                            for index, result in enumerate(
                                online_results,
                                start=1
                            ):

                                title = result.get(
                                    "title",
                                    "Source"
                                )

                                url = result.get("href", "")

                                if url:
                                    answer += (
                                        f"{index}. "
                                        f"[{title}]({url})\n\n"
                                    )

                        else:

                            answer += (
                                "\n\nOnline search was unavailable "
                                "for this request. This explanation "
                                "uses available study material and "
                                "general knowledge."
                            )

                        return answer

                    errors.append(
                        f"{model_name}: Empty response"
                    )

                    break

                except Exception as error:

                    error_text = str(error)

                    errors.append(
                        f"{model_name}: {error_text}"
                    )

                    if attempt == 0 and any(
                        code in error_text
                        for code in [
                            "503",
                            "UNAVAILABLE",
                            "429",
                            "RESOURCE_EXHAUSTED"
                        ]
                    ):

                        time.sleep(2)
                        continue

                    break

        return (
            "All Gemini models are temporarily unavailable.\n\n"
            + "\n".join(errors)
        )

    except Exception as error:

        return f"AI generation failed: {error}"
