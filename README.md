# xlsxt

Transform xlsx template to xlsx

## How it works

providing an excel as template

![image](https://github.com/user-attachments/assets/158ceda7-e569-44ca-8369-1c9268d1bf22)

and some nested/structured data

![image](https://github.com/user-attachments/assets/5d7932f6-f7e3-48ef-b13a-8d2876794708)

you get an instantiated template

![image](https://github.com/user-attachments/assets/3f1217b1-5e19-4551-a56a-3c31ea1cef5d)


## Installing and running

```
uv venv --python 3.11
source ./.venv/bin/activate
uv pip install -r requirements.txt
uv run python demo.py
# check output.xlsx will use demo.json
uv run python demo.py loadtest
# check output.xlsx will use a much larger generated context
```

## Todo

- [ ] make it a proper python package, that can be published or at least used from git pip install
- [ ] support other cell formats
   - [ ] date
   - [x] hyperlink (notation `https://...|label to be shown`)
   - [ ] ... (I don't know much about excel)
- [ ] make it a cli taking xlsx and json (file or stdin) and output file destination
- [ ] investigate better error handling
   - [ ] (more context source line in the template, cell coordinates, formulas, perhaps a did you mean ?)
   - [ ] a more "lenient" mode where the cell is colored red and extras sheets with errors info ?
- [ ] investigate easier formula authoring for "sum" and current row calculations
   - [ ] something where you really use excel formulas and subtitute/transpose the ranges
        - [x] for same line
        - [ ] for a vertical group of items : would be great for sub totals
        - goal : get autocompletion and validation from the formula editor while authoring, not when previewing
- [ ] document post processing and add some extra post processing
   - [ ] hide a sheet
   - [ ] delete a sheet (ex used for config )
- [ ] play with wasm and offer a live preview ? or allow the sheet to be in googlesheet ?
- [ ] have a better code base and setup
  - [ ] currently a huge file without too much thought
  - [ ] investigate a way to unit/integration test the damn thing : see comparator
  - [ ] github actions
      
## Inspirations

this project is heavily inspired by https://github.com/ivahaev/go-xlsx-templater

other similar projects 
  - xlsx-template
    - https://github.com/optilude/xlsx-template/
    - js 
    - looks limited in term of tables/nesting
  - docxtemplater
    - https://docxtemplater.com/modules/xlsx/#limitations 
    - js commercial
  - xlsx-templater
    - https://github.com/yangguichun/xlsx-templater
    - js
    - looks compatible with some of the "commercial" 
  - xltpl
    - https://github.com/zhangyu836/xltpl
    - https://pypi.org/project/xltpl/#history
    - python
    - seem to support more "jinja" syntax like if elif
      - ![image](https://github.com/user-attachments/assets/5e0a5a8f-a7bf-42d8-b131-95161cd117fd)
    - looks okish but haven't tried, 
      - no real test suite, only examples
      - comments/docs in chinese (I guess)

note most have limitation in "translating" formulas, which is one the thing I struggle with to avoid "string concatenation".

in the more advanced things I see they talk about data validations in excel
