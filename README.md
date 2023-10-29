# Create PowerPoint for worship lyrics

## File Format

### Starting Line

Every file must start with this line:

```text
!VER 2
```

This is the version indicator to avoid incompatible versions of this program in processing the lyrics file.

### Background image

The background image of furture slides can be defined with:

```text
!BKG <path>
```

Where `<path>` is the path to the image, either absolute or relative to the lyrics file.

### Fonts

Both lyrics and footers can be customized. The following can be customized:

* `FONT`: The name of the font to be used. It is recommended to include a font that is known to exist on both your computer and the one presenting the PowerPoint.
* `SIZE`: Font size in Pt. This should be a valid floating point number.
* `COLOR`: 8-bit hexadecimal RGB color. For example, `000000` represents pure black.

Commands in the format of `!<scope>-<type>`, where `<scope>` can be either `LYRICS` or `FOOTER`, define the format of texts in that scope occurring after this statement. Note that the occurring time of the footer is the time a new slide is created, not the time the footer is defined.

Commands in the format of `!<type>` imply the command of both scopes.

For example, the following snippet defines the font for both lyrics and footers, then defines different colors and sizes for each:

```text
!FONT Arial
!LYRICS-COLOR FFFFFF
!FOOTER-COLOR 000000
!LYRICS-SIZE 60
!FOOTER-SIZE 24
```

### Lyrics Text

A block of lyric texts is separated from other lyric texts by two spaces, or in other words, there is one empty line between two lyric text blocks. For example, the following snippet represents two lyric text blocks:

```text
My table thou hast furnished
in presence of my foes:

my head thou dost with oil anoint,
and my cup overflows.
```

Each lyric text blocks will be put into individual slides.

### Sections

Sections are wrapped with `!SECTION-START <section-name>` and `SECTION-END`. For example, the following snippet creates a section called "`Shep-C`" with two lyric text blocks:

```text
!SECTION-START Shep-C
Goodness and mercy all my life
shall surely follow me.

And in God's house for evermore
my dwelling place shall be. 
!SECTION-END
```

A section does not do anything on its own. They can be used via `!SECTION <section-name>`. For example, the following snippet repeats the above section twice:

```text
!SECTION Shep-C
!SECTION Shep-C
```

Note that a section must be defined before their usage, and cannot be defined more than once using the same name.

### Indication of empty slides

An indication of an empty slide is declared as:

```text
!EMPTY
```

This creates a footer-only slide without any lyrics.

## Examples

Examples of this program can be found in `examples/`. Unless otherwise specified, all lyrics and images inside are in the Public Domain.
