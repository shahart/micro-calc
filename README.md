# micro-calc
modified 'excel' for J2ME to support PMT financial function, and export

Refer to http://boostapps.com/apps/microcalc/

Upload its Jar into https://github.com/zb3/freej2me-web

--

Asked Codex to re-build the MIDlet with my code changes:

```
javac --release 8 -classpath '/tmp/midpapi20-2.0.4.jar:/tmp/cldcapi11-2.0.4.jar:microcalc.jar' \
  -d /tmp/microcalc-rebuild/classes \
  /tmp/microcalc-rebuild/mathsrc/net/jscience/math/MathFP.java \
  BadFormulaException.java Menu.java MicroCalc.java MicroCalcCanvas.java \
  Result.java Sheet.java TempScreen.java Util.java

jar cfm microcalc-rebuilt.jar /tmp/microcalc-rebuild/MANIFEST.MF \
  -C /tmp/microcalc-rebuild/classes .
```

The core interaction is a 32×8 sheet with formula editing, a menu for cell/sheet/row/column actions, and SYLK export (see https://en.wikipedia.org/wiki/Symbolic_Link_(SYLK))

Now asked to impl' those workflows in a responsive browser UI, with keyboard navigation, local sheet storage, JSON backup, and the formula/function set from the source.

`node --test tests/calc-core.test.mjs` cover arithmetic, references and ranges, aggregate functions, PMT, logical and string formulas, dates, formula errors, ...

plus more gaps: first-key entry, row/column edits and reference updates, and workbook persistence/import/export.

For details why I did it, ask AI: `explain me this mortgage calc: =PMT(4.1/12/100,12*20,1000*130)`

## Static site TypeScript source

The GitHub Pages site is in `docs/`. TypeScript sources are in `docs/src/`; `npm run build` compiles them to the browser modules in `docs/dist/`.

```sh
npm install
npm run build
npm test
```
