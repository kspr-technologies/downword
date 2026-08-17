The transformer architecture[^1] replaced recurrent models[^rnn] for most
sequence tasks, though RNNs are still used in some streaming settings[^rnn].

An inline footnote works too^[defined right here, with no label].

A reference with no definition stays as literal text: [^missing].

[^1]: Vaswani et al., *Attention Is All You Need* (2017).

[^rnn]: Recurrent neural networks.

    A second paragraph inside the same footnote.

    - and a list
    - inside it

    ```py
    and_even_code()
    ```

[^unused]: This definition is never referenced, so markdown-it-footnote drops it.
