# 📝 PDF → Excel (OCR)

Este script converte um PDF em planilha Excel formatada, extraindo os campos:
- ITEM
- CATMAT
- DESCRIÇÃO DETALHADA
- UNIDADE
- QUANTIDADE
- VALOR UNITÁRIO
- VALOR TOTAL

## 🚀 Como usar

### 1. Clone o repositório
```bash
git clone https://github.com/seuusuario/seurepo.git
cd seurepo
```

### 2. Crie um ambiente e instale as dependências
```bash
pip install -r requirements.txt
```

⚠️ Dependências do sistema:
- **poppler-utils** → necessário para `pdf2image`
- **tesseract-ocr-por** → OCR em português

No Linux/WSL:
```bash
sudo apt-get install -y poppler-utils tesseract-ocr-por
```

### 3. Execute o script
```bash
python script.py entrada.pdf -o saida.xlsx
```
