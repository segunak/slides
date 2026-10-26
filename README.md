# Slides

<https://segunakinyemi.com/slides/>

A place for my slide decks because everyone these days has started making web-based decks because it's easy with AI agents.

But I don't want to give up manual editing in some form. I don't want a totally AI dependent (tokens aren't free) workflow for editing slide decks. I still want to be able to click a slide, change some text, and move things around, as I've grown accustomed to in PowerPoint/. So I have this repo, which uses [PPTX2HTML](https://github.com/g21589/PPTX2HTML) to convert my PowerPoints into HTML web endpoints I can easily share. This repo includes automation to watch my OneDrive for certain decks and automatically update the web endpoint when I edit decks.

[Pamela Fox's ai-powered-presentation-workflow](https://github.com/pamelafox/ai-powered-presentation-workflow) was the catalyst for me setting this up. She's able to have amazing slides like <https://pamelafox.github.io/my-py-talks/impact-of-ai-on-cs/>, backed by a GitHub repo <https://github.com/pamelafox/my-py-talks/tree/main/impact-of-ai-on-cs>, fully edited by agents as she explains in <https://pamelafox.github.io/ai-powered-presentation-workflow/>, and I was like, I want that.

But I also want to at least keep the "hey, I can still manually edit my slides from time to time" piece. I've grown comfortable with PowerPoint, and I think the limitations agents had in the past editing PowerPoints are totally gone. I use Anthropic's [PowerPoint agent skill](https://github.com/anthropics/skills/tree/main/skills/pptx) to edit decks, and it's amazing, fully able to work with PowerPoint and make whatever you want, and you can still go in and set the tone, provide some direction, and choose your own theme/style, and it just goes off that.

What I don't want is to find myself needing to edit a slide deck in a hurry and it becomes "oh shoot got to spin up an agent session" because the HTML it produced is so convoluted I can hardly parse it without help. Keeping some sort of GUI based experience has tremendous value to me.

So, yeah.
