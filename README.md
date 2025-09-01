# PowerPoint Notes
This obsidian plugin allows to synchronize your PowerPoint slides with a markdown note in Obsidian.

Currently, it only supports the direction to overwrite notes in the PowerPoint file with the contents of the markdown note.
The slide content is not changed, only the presentation notes are updated.

Each header corresponds to a slide in the PowerPoint presentation, and the bullet points under each header are used as the notes for that slide.

Example: 

```md
---
powerPoint-file: folder\presentation.pptx
---
# Slide 1
- Note 1
- Note 2

# The title of the Header is not important
- Note of Slide 2
    - Nesting is supported
```

To specify the PowerPoint file to update, you need to add a frontmatter property with the name `powerPoint-file` and the path to the PowerPoint file.

When you want to write your notes to the PowerPoint file, ensure that PowerPoint is closed and that you have a backup copy of your presentation, as I only tested my own files and I can't guarantee it will work with yours.

You can use the command `PowerPoint Notes: Write Notes to Slides` to trigger the note writing process.