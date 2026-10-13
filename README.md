# HCC Group website

Website for the Human-Centered Computing Group at Tulane University, co-directed by
Rebecca Faust, Saad Hassan, and Brian Summa.

It is one static page (`index.html`) generated from a spreadsheet. There is no framework
and nothing to install beyond Python and `openpyxl`.

## Updating content

1. Open `content/hcc-content.xlsx` and edit the tab you need. The README tab in the
   spreadsheet explains each column.
2. Run the build:

   ```sh
   pip install openpyxl   # once
   python3 build.py
   ```

3. Open `index.html` in a browser (or run `python3 -m http.server` and visit
   http://localhost:8000) to check it, then commit and push.

If the GitHub workflow is enabled, pushing also rebuilds and deploys the site, so step 2 is
optional for a quick fix.

### Common edits

| To do this | Edit |
| --- | --- |
| Add news | New row in **News**. Put `Yes` in Highlight? to pin it at the top (up to six). |
| Add a paper | New row in **Papers**. Set Project to a project key. `Yes` in Highlight? adds it to Highlighted. |
| Add a person | New row in **Students** (or **FacultyStaff**). Put the photo in `images/people/` and its path in Photo. |
| Move a graduate to alumni | Set Status to `Former`, fill Last year in lab and Next position. |
| Add a grant | New row in **Grants**. Grants show largest first. Tulane internal awards and anything under the Settings minimum ($10,000) are left off. |
| Change a project | Edit **Projects**. Projects appear alphabetically by title. |
| Link a co-author's name | Add them to **Coauthors**. Group members are linked automatically. |

Photos: square, about 320 × 320 px, named `firstname-lastname.jpg`. Everyone current with a photo also
appears in the photo grid at the top of the page, so new members show up there automatically. Leave Photo alt text
empty for a plain portrait (the name sits right next to it) or describe the photo.

Dates in News can be a full date or just a year.

## Files

```
index.html                    generated page (do not edit by hand)
build.py                      spreadsheet → index.html
templates/index.template.html page structure, header, footer, section intros
content/hcc-content.xlsx      all content
css/site.css                  styles; colors are tokens at the top, with dark mode
js/site.js                    theme switch, menu, news and publication filters
fonts/                        Bricolage Grotesque and Inclusive Sans (SIL OFL)
images/                       people, paper thumbnails, favicon
.github/workflows/deploy.yml  builds and deploys to GitHub Pages
```

## Accessibility

The page targets WCAG 2.2 AA with semantic landmarks and headings, a skip link, visible focus,
reduced-motion support, light and dark themes that follow the system setting, and filters
that announce their results. Everything still works with JavaScript turned off. Check with
axe or a screen reader after larger changes.
