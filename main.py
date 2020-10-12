from flask import Flask
from flask import request
from pptx import Presentation

app = Flask(__name__)


def create_ppt(ppt_title):
    print("in create_ppt")
    prs = Presentation()
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]

    title.text = ppt_title
    subtitle.text = "Created with Powerpoint Bot"

    prs.save(ppt_title + ".pptx")
    return


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
    app.run(debug=True)  # run server in debug mode
