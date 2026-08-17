A fence with a language and an extra info string:

```ts title=foo.ts
export const answer = 42;
```

A fence with a language, a filename and a line-highlight range:

```js title="src/app.js" {1,3-5} showLineNumbers
const a = 1;
```

A fence with no language at all:

```
plain text, no highlighting
```

A fence with only whitespace in the info string:

```   
still no language
```

A fence using tildes:

~~~diff
- old line
+ new line
~~~

A four-backtick fence containing a three-backtick fence:

````markdown
```js
nested();
```
````

An indented code block (four spaces), which has no info string at all:

    indented();
    also_indented();

A fence whose language is oddly cased and dotted:

```C#
var x = 1;
```
