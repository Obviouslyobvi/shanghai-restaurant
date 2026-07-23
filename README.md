# ShangHai 上海小馆 — spec site

Spec website for ShangHai Restaurant, 830 N Milpas St, Santa Barbara, CA 93103.
Static. No build step, no dependencies, no server code.

## Live

| URL | Design |
|---|---|
| `/` | Bright — enamelware palette, moon gate arches, Bricolage Grotesque |
| `/v1/` | Dark — Shanghai Deco, lacquer and brass, Bodoni Moda |
| `/v2/` | Same as `/` (kept so the original v2 link still resolves) |

## Files

```
index.html      bright design (homepage)
v1/index.html   dark design, images referenced as ../img/
v2/index.html   copy of the bright design
img/            hero + 4 dish photos, webp
v2/img/         same images, for the /v2/ path
```

Fonts load from Google Fonts. Everything else is self-contained.

## Before this goes to the public

- Replace delivery-app prices with the counter price list from the owner
- Confirm the uDinner deep link lands on their live order page
- Swap the Grubhub neighborhood link for their store URL
- Confirm which photo goes with which dish
- Add lunch specials once the owner supplies the current list

## Editing the menu

All 172 dishes live in the `const MENU = [...]` array near the bottom of each
`index.html`. Each item is `{n: name, p: price, d: description, v: vegetarian,
s: spicy}`. Change a price by editing `p`. Nothing else needs touching.
