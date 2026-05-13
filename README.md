# App web - Consolidado de Emendas

Este app cria uma interface web para os dois fluxos:

1. Consolidado completo: recebe um ZIP com todas as emendas em PDF e gera `Consolidado_Emendas.xlsx`.
2. Consolidado Incremental: recebe `Emendas.xlsx` e, opcionalmente, um ZIP com novas emendas em PDF. Preserva a aba `Emendas` enviada, acrescenta PDFs ainda não listados quando houver, e gera `Consolidado_Incremental.xlsx`.

## Rodar localmente para teste

```powershell
pip install -r requirements-web.txt
python web_app.py
```

Acesse: `http://localhost:8000`

## Publicar para colegas sem instalação local

Use o `Dockerfile` em um serviço como Render, Railway, Azure App Service, Google Cloud Run ou servidor interno. O container já instala Tesseract OCR com idioma português.

Variáveis úteis:

- `MAX_UPLOAD_MB`: limite do upload em MB. Padrão: `900`.
- `SECRET_KEY`: chave usada pelo Flask em produção.

Observação: OCR pode demorar bastante em arquivos grandes. Use plano/servidor com timeout alto e pelo menos 2 GB de RAM.
