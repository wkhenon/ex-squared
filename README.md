# Ex Squared

Ex Squared turns a selected expense section in an Obsidian note into a formatted Excel summary.

## Export expenses

1. In a Markdown note, select from the `Budget for Period` heading through the last expense bullet in the period.
2. Run **Export selected expenses to Excel** from the Command palette, or right-click the selection and choose the same command.
3. The plugin saves a date-stamped `.xlsx` file in the vault's `Exports` folder and inserts a link to it after the selected section. If that filename already exists, it adds a number to avoid overwriting the earlier file.

The workbook groups expenses by card. Each card has a section header with its total, followed by payment-source subtotals. The first hashtag on each expense bullet identifies the payment source; the second identifies the card. Day-level parenthetical totals and other day-level tags are ignored.

For example:

```markdown
***Budget for Period: $0.00***
- 01/01 monday #example-day-tag
  - #example-source #example-card $0.00 sample expense
```

The plugin processes notes locally and does not transmit note contents.

## License

Ex Squared is licensed under GPL-3.0-only. The production bundle includes fflate, which is licensed under MIT; its license notice is included in the generated `main.js` file.

## Development

- `npm install` installs dependencies.
- `npm run dev` watches and bundles the plugin.
- `npm run build` type-checks and creates the production bundle.
