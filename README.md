# Services

Collection of independent, publicly-hosted utility tools. Each service lives in its own folder and is fully self-contained (HTML + assets, no build step).

## Live URLs

- Index: [rayhanul17.github.io/services](https://rayhanul17.github.io/services/)
- Tower of Hanoi: [rayhanul17.github.io/services/learn_th](https://rayhanul17.github.io/services/learn_th/)
- Voucher: [rayhanul17.github.io/services/voucher](https://rayhanul17.github.io/services/voucher/)
- MessMate: [rayhanul17.github.io/services/messmate](https://rayhanul17.github.io/services/messmate/)

## Adding a new service

1. Create a new folder at the repo root, e.g. `services/invoice/`.
2. Add `index.html` and any assets inside it.
3. Add a card to the root `index.html` linking to it.
4. (Optional) Develop on a feature branch (`invoice-dev`) and merge to `main` to deploy.

<!-- markdownlint-disable MD060 -->
## Available services

| Service | Folder                  | Description                                         |
|---------|-------------------------|-----------------------------------------------------|
| Tower of Hanoi | [learn_th/](learn_th/) | Interactive Tower of Hanoi learning game and solver |
| Voucher | [voucher/](voucher/)    | মেসার্স অর্পা কনস্ট্রাকশন site voucher generator    |
| MessMate | [messmate/](messmate/) | Local-first mess management, settlements, and reports |
<!-- markdownlint-enable MD060 -->
