@echo OFF

call C:\Users\phbin\anaconda3\Scripts\activate.bat C:\Users\phbin\anaconda3

python quizizz-template-generate.py

call conda deactivate