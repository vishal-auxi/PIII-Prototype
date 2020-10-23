from flask import Flask, request, jsonify
from pptx import Presentation
import wikipedia

from Summarizer import summarize_lsa, sentence_count
from team_slide import create_team_slide
from pie_chart import create_pie_chart
from bar_chart import create_bar_chart

app = Flask(__name__)


def get_summarized_sections(ppt_subject):
    pages_search = wikipedia.search(ppt_subject, results=5, suggestion=False)
    page_suggest = wikipedia.suggest(ppt_subject)

    if not pages_search and not page_suggest:
        return False

    page_match = ppt_subject
    if pages_search:
        print(pages_search)
        for page in pages_search:
            print("in loop for i = " + page)
            try:
                wikipedia.page(page).content
                page_match = page
                break
            except wikipedia.DisambiguationError as e:
                continue
            except Exception as e:
                continue
        # page_match = pages_search[0]
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
        return (False,)

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

    prs.save("Presentation - " + page_title + ".pptx")

    return True, page_title


def create_ppt_with_count(ppt_subject, count):
    print("in create_ppt_with_count")

    result = get_summarized_sections(ppt_subject)
    if not result:
        return (False,)

    page_title = result[0]
    sections = result[1]
    url = result[2]

    prs = Presentation()
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]

    title.text = page_title
    subtitle.text = "Created by Presentation Bot \n (Sourced from {url})".format(url=url)

    total = 0
    for section in sections.items():
        if total >= count:
            break

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

        total += 1

    prs.save("Presentation - " + page_title + ".pptx")

    return True, page_title, total


@app.route("/", methods=['POST'])
def ppt_request_handle():
    print("in ppt_request_handle")
    req = request.json
    print(req)

    if req["req"] == "create_ppt":
        result = create_ppt(req["title"])
        if result[0]:
            response = {
                "message": "Request successful",
                "title": result[1]
            }
            return jsonify(response), 200

    elif req["req"] == "create_ppt_count":
        result = create_ppt_with_count(req["title"], req["count"])
        if result[0]:
            response = {
                "message": "Request successful",
                "title": result[1],
                "count": result[2],
            }
            return jsonify(response), 200

    elif req["req"] == "create_team_slide":
        result = create_team_slide(req["people"])
        response = {
            "message": result[1]
        }
        if result[0]:
            return jsonify(response), 200
        else:
            return jsonify(response), 400

    elif req["req"] == "create_pie_chart":
        result = create_pie_chart(req)
        response = {
            "message": result[1]
        }
        if result[0]:
            return jsonify(response), 200
        else:
            return jsonify(response), 400

    elif req["req"] == "create_bar_chart":
        result = create_bar_chart(req)
        response = {
            "message": result[1]
        }
        if result[0]:
            return jsonify(response), 200
        else:
            return jsonify(response), 400

    return f"Cannot process request", 400


# 'Main' function to run
if __name__ == '__main__':
    app.run(debug=True)  # run server in debug mode

    # create_team_slide(["chris", "john", "mike", "steve"])
    # res = create_pie_chart({
    #     'req': 'create_pie_chart',
    #     'categories': ['USA', 'Canada', 'Mexico'],
    #     'percentages': ['30', '30', '40']}
    # )

    # res = create_bar_chart({
    #     'req': 'create_bar_chart',
    #     'categories': ['USA', 'Canada', 'Mexico'],
    #     'values': ['81', '45', '54']}
    # )

    # print(res)
