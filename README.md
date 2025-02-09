# xlsxt

Transform xlsx template to xlsx

## How it works

providing an excel as template

![image](https://github.com/user-attachments/assets/158ceda7-e569-44ca-8369-1c9268d1bf22)

and some nested/structured data

![image](https://github.com/user-attachments/assets/5d7932f6-f7e3-48ef-b13a-8d2876794708)

you can an instantiated template

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
