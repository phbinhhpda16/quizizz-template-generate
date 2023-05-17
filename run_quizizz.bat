@echo OFF

call C:\...\anaconda3\Scripts\activate.bat C:\...\anaconda3

python quizizz-template-generate.py

call conda deactivate
