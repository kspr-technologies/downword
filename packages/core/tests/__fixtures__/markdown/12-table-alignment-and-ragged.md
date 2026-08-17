| Left | Center | Right | Default |
|:---- |:------:| -----:| ------- |
| a | b | c | d |
| 1 | 2 | 3 | 4 |

Ragged rows — markdown-it pads the short one and truncates the long one:

| one | two | three |
| --- | --- | ----- |
| only one cell |
| a | b | c | d | e |

A header-only table with no body rows:

| Column A | Column B |
| -------- | -------- |

A table where the delimiter row is misaligned but still valid:

|x|y|
|-|-|
|1|2|
