#cd ~/bin/QtDesigner/
#source ~/bin/QtDesigner/.venv/bin/activate

# IO table generator
pyuic6 -o iogen_main.tmp ui/iogen_main.ui
sed 's/:\/asset\//asset:/g' iogen_main.tmp > iogen_main.py
rm iogen_main.tmp

#rcc -g python -o resources.py asset/res.qrc