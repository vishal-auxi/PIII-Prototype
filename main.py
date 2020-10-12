from flask import Flask, request
from pptx import Presentation
import wikipedia

from Summarizer import summarize_lsa, sentence_count

app = Flask(__name__)


def get_summarized_sections(ppt_subject):
    print("in get_summarized_sections")

    pages_search = wikipedia.search(ppt_subject, results=5, suggestion=False)
    page_suggest = wikipedia.suggest(ppt_subject)

    if not pages_search and not page_suggest:
        return False

    page_match = ppt_subject
    if pages_search:
        page_match = pages_search[0]
    elif page_suggest:
        page_match = page_suggest

    page = wikipedia.page(page_match)

    title = page.title
    url = page.url
    summary = page.summary
    images = page.images
    categories = page.categories
    content = page.content
    links = page.links
    sections = page.sections
    references = page.references

    if not sections:
        return False

    sections_summarized = {}
    for section in sections:
        if section.lower() == "see also":
            break
        sections_summarized[section] = summarize_lsa(page.section(section))

    return title, sections_summarized, url


def create_ppt(ppt_subject):
    print("in create_ppt")

    result = get_summarized_sections(ppt_subject)
    if not result:
        return False

    page_title = result[0]
    sections = result[1]
    url = result[2]

    # for section in sections.items():
    #     print(section[0] + ":")
    #     print(section[1])
    #     print("\n")

    prs = Presentation()
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]

    title.text = page_title
    subtitle.text = "Created by Presentation Bot \n (Sourced from {url})".format(url=url)

    for section in sections.items():
        if not section[1]:
            continue

        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        subtitle = slide.placeholders[1]

        title.text = section[0]
        subtitle_text = ""
        for i in section[1]:
            subtitle_text = subtitle_text + "\n" + i
        # subtitle.text = section[1][0]
        subtitle.text = subtitle_text

    prs.save(page_title + " - Presentation.pptx")

    return


# def create_ppt(ppt_title):
#     print("in create_ppt")
#     prs = Presentation()
#     title_slide_layout = prs.slide_layouts[0]
#     slide = prs.slides.add_slide(title_slide_layout)
#     title = slide.shapes.title
#     subtitle = slide.placeholders[1]
#
#     title.text = ppt_title
#     subtitle.text = "Created with Powerpoint Bot"
#
#     prs.save(ppt_title + ".pptx")
#     return


@app.route("/", methods=['POST'])
def ppt_request_handle():
    print("in ppt_request_handle")
    req = request.json
    print(req)

    if req["req"] == "create_ppt":
        create_ppt(req["title"])
    return f"Request successful", 200


# 'Main' function to run
if __name__ == '__main__':
    # app.run(debug=True)  # run server in debug mode
    create_ppt("Coronavirus")
