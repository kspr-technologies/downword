The model got cut off mid-answer, which happens constantly:

Here is the corrected implementation:

```python
def process(items):
    results = []
    for item in items:
        if item.is_valid():
            results.append(transform(item))
    return results
