# How to work
![Flow](image-1.png)


## setup
- git clone https://github.com/niknew1996/diggy_stand_local.git
- cd diggy_stand_alone
- install python from this https://www.python.org/downloads/
- pip install virtualenv
    1. python -m venv myenv  # สร้าง virtual environment
    2. source myenv/Scripts/activate  # เปิดใช้งาน virtual environment
    3. pip install requirements.txt
- เปิด Excel แล้วใส่ Source ip start-end / Destinations ip start-end / port / user - password สำหรับ ssh
## Run
- python test_ssh.py
- จะถามช