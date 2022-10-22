run: create_venv
	./.venv/Script/active 
	python app.py

create_venv: requirements.txt
	python -m venv .venv
	./venv/bin/pip install -r requirements.txt