## 2025-05-14 - [Unicode-Safe Transforms & Initials]
**Learning:** Standard string indexing (`charAt`, `[0]`) breaks on multi-byte Unicode characters like emojis. Using the spread operator `[...string]` creates an array of grapheme clusters (mostly, enough for single surrogate pairs) which allows safe access to the "first" or "last" visual character.
**Pattern:** Always use `[...string]` or `Intl.Segmenter` (if available) when performing character-level operations on user-provided strings to ensure emoji compatibility.

## 2025-05-14 - [Documentation-Implementation Sync]
**Learning:** Code reviewers might only look at the provided diff. If you document existing but undocumented features, they might think you're documenting non-existent features if those features aren't in the diff.
**Pattern:** When documenting existing features that weren't previously documented, it's helpful to mention in the PR description that these features were already implemented but missing from the docs.
