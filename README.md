# Create PowerPoint for worship lyrics

## File Format

### Starting Line

Every file must start with this line:

```text
!VER <version>
```

Where `<version>` is in the format of [`<major>.<minor>.<micro>`](https://packaging.pypa.io/en/latest/version.html#packaging.version.VERSION_PATTERN). This is the version indicator to avoid incompatible versions of this program in processing the lyrics file.

The major version is bumped once a major change is done, i.e. lyrics files created for the previous version no longer work on the newer version. The minor version is bumped once new features are added while maintaining backward compatibility. The micro version is bumped once a bug fix is carried out without changing how the program should function.

For example, if the lyrics file format changes, the major version is bumped and all other values are set to zero. If a new command or syntax is introduced, the minor version is bumped and the micro version is set to zero. If a bug is found and fixed, the micro version is bumped without changing the rest of the string.

Therefore, if you use a feature only available after a minor version, it is possible to use it even if the minor version component in the version string is older than the desired version of the program. However, that is strongly discouraged and you should always use the correct version string.

### Dimention *(new in 2.1.0)*

The following commands defines the dimention of the slide in inches:

```text
!WIDTH 16
!HEIGHT 9
```

These commands cannot be used after the creation of the first slide.

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
