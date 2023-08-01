# no human lookup
귀찮아서 더 귀찮은 일 만들기
<br/><br/><br/>

### 🔨 환경 구성 (vscode)
- `` Ctrl + ` `` (open terminal)
- `$ python -m venv venv` (create virtual environment named *venv*)
- `$ source venv/Scripts/activate` (activate virtual environment)
- `$ pip install -r requirements.txt` (install required dependencies)
- `$ python main.py` (run)


<br/>


### ⚙ DB connection 정보 세팅

create `db.properties` file on root directory and set db properties 

ex)
```properties
# db credentials
user=root
passwd=mypassword1234
host=localhost
db=myschema
charset=utf8
```

<br/>

