ls
sudo apt update && sudo apt upgrade -y
sudo apt install python3-pip python3-venv build-essential libsndfile1 -y
python3 -m venv venv
source venv/bin/activate
pip install flask flask_sqlalchemy flask_cors flask_socketio apscheduler soundfile scipy numpy pytz
mkdir templates static uploads
ls
touch app.py
nano app.py
ls
cd templates
ls
touch analysis.html
ls
nano analysis.html
touch comparacion.html
nano comparacion.html
ls
touch error.html
nano error.html
touch index.html
nano index.html
touch monitoring.html
nano monitoring.html
touch monitoring_finished.html
nano monitoring_finished.html
ls
cd ..
ls
source venv/bin/activate
nohup python3 app.py > output.log 2>&1 &
cd ~
mkdir nueva-app && cd nueva-app
python3 -m venv venv
source venv/bin/activate
pip install flask flask_sqlalchemy flask_cors flask_socketio apscheduler soundfile scipy numpy pytz
touch app.py
nano app.py
nohup venv/bin/python3 app.py > nueva.log 2>&1 &
mkdir templates static uploads
ls
,ls
ls
source venv/bin/activate
nohup python3 app.py > output.log 2>&1 &
ls
nano app.py
cd tempaltes
ls
cd templates
ls
nano index.html
ls
cd ..
ls
sc static
cd static
ls
cd ..
nano app.py
ls
cd instance
ls
cd templates
ls
cd app.py
ls
cd templates
ls
rm monitoring_finished.html
touch monitoring_finished.html
nano monitoring_finished.html
cd templates
rm monitoring_finished.html
touch monitoring_finished.html
nano monitoring_finished.html
cd templates
rm monitoring_finished.html
touch monitoring_finished.html
nano monitoring_finished.html
cd templates
ls
cd monitoring.html
nano analysis.html
cd templates
rm monitoring_finished.html
cd templates
touch monitoring_finished.html
nano monitoring_finished.html
source venv/bin/activate
nohup python3 app.py > output.log 2>&1 &
ls
nano app.py
ps aux | grep app.py
curl http://checkip.amazonaws.com
