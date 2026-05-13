# Deploy na Oracle Cloud Always Free

## VM recomendada

- Image: Ubuntu 22.04 ou Ubuntu 24.04
- Shape: `VM.Standard.A1.Flex`
- OCPU/RAM: comece com `2 OCPU` e `12 GB RAM`
- Rede: VCN publica com IP publico

Evite `VM.Standard.E2.1.Micro` para este app, porque OCR e PDFs grandes exigem mais CPU e memoria.

## Portas

Libere a porta `8000` em dois lugares:

1. Oracle Cloud: Security List ou Network Security Group da VCN.
2. Ubuntu, se o firewall estiver ativo:

```bash
sudo ufw allow 8000/tcp
sudo ufw reload
```

## Instalar e subir o app

Conecte na VM por SSH e rode:

```bash
curl -fsSL https://raw.githubusercontent.com/victorramon88-dev/painel-emendas/main/deploy_oracle.sh -o deploy_oracle.sh
chmod +x deploy_oracle.sh
./deploy_oracle.sh
```

O app ficara em:

```text
http://IP_PUBLICO_DA_VM:8000
```

## Atualizar depois

Quando houver nova versao no GitHub:

```bash
cd ~/painel-emendas
git pull --ff-only
sudo docker compose up -d --build
```

## Logs

```bash
cd ~/painel-emendas
sudo docker compose logs -f
```
