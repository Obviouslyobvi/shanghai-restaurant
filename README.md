# ShangHai 上海小馆 — spec site

Spec website for ShangHai Restaurant, 830 N Milpas St, Santa Barbara, CA 93103.
Static. No build step, no dependencies, no server code.

## Files

```
shanghai/
├── index.html      52 KB   entire site: markup, CSS, JS, menu data
└── img/           139 KB
    ├── hero.webp
    ├── d1.webp    wor wonton soup
    ├── d2.webp    soup dumplings
    ├── d3.webp    green beans / chow mein
    └── d4.webp    grilled whole fish
```

Total 208 KB. Loads Bodoni Moda, Manrope, and Noto Serif SC from Google Fonts.

## Deploy to GitHub Pages (web UI, no terminal)

1. Open `github.com/Obviouslyobvi/obviouslyobvi.github.io`
2. **Add file → Upload files**
3. Drag the `shanghai` folder in whole
4. Commit to `main`
5. Live in about a minute at `obviouslyobvi.github.io/shanghai/`

## Deploy from a terminal instead

```bash
cd path/to/obviouslyobvi.github.io
cp -r shanghai .
git add shanghai && git commit -m "ShangHai spec site" && git push
```

## Before this goes to the public

- Replace delivery-app prices with the counter price list from the owner
- Confirm the uDinner deep link lands on their live order page
- Swap the Grubhub neighborhood link for their store URL
- Confirm which photo goes with which dish
- Add lunch specials once the owner supplies the current list

## Editing the menu

All 172 dishes live in the `const MENU = [...]` array near the bottom of
`index.html`. Each item is `{n: name, p: price, d: description, v: vegetarian,
s: spicy}`. Change a price by editing `p`. Nothing else needs touching.
