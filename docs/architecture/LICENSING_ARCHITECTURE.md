# Arquitetura de licenciamento luGEST

## Objetivo

Preparar um licenciamento comercial verificável sem misturar regras comerciais,
interface, base de dados do cliente e segredos de assinatura.

O código criado nesta fase é uma fundação isolada. O trial atual continua ativo
e nenhuma licença comercial passa a bloquear a aplicação até serem aprovadas as
regras de negócio.

## Limite de confiança

O programa instalado no cliente recebe apenas uma **chave pública**. A chave
privada usada para emitir licenças nunca entra no repositório, no executável, no
instalador ou na infraestrutura do cliente.

```text
Ferramenta interna de emissão          Instalação do cliente
--------------------------------       --------------------------------
chave privada (cofre/serviço)          chave pública
        |                                      |
        v                                      v
payload + assinatura  ------------->   valida assinatura e regras
                                               |
                                               v
                                      permite ou bloqueia módulos
```

## Componentes

- `lugest_core/licensing/license.py`: formato, validação criptográfica e política
  pura de validade, equipamento, módulos e postos.
- `lugest_core/licensing/trusted_time.py`: hora HTTPS independente do relógio do
  Windows, já usada pelo trial.
- `lugest_infra/licensing/license_store.py`: gravação atómica do token fora dos
  dados operacionais e da configuração editável da UI.
- `scripts/verify_license_foundation.py`: prova automatizada de assinatura,
  adulteração, expiração, vínculo a equipamento, módulos, postos e persistência.

## Formato v1

O envelope é `LUGEST1.payload.assinatura`, com Base64 URL-safe. O payload JSON é
canónico e contém:

- identificador da licença e do cliente;
- nome do cliente e edição;
- emissão, início e fim de validade em UTC;
- limite de postos e utilizadores;
- módulos autorizados;
- equipamentos autorizados, quando aplicável;
- metadados comerciais não secretos.

A assinatura prevista é Ed25519. Alterar um único byte invalida o token.

## Decisões que faltam antes da integração

1. Definir edições comerciais e módulos incluídos em cada uma.
2. Decidir se o limite é por posto instalado, sessão simultânea ou utilizador.
3. Definir tolerância offline e frequência de renovação/validação online.
4. Definir transferência de licença quando o cliente muda de computador.
5. Definir período de graça para falhas de Internet ou indisponibilidade do
   serviço de licenciamento.
6. Definir fluxo de revogação, renovação, auditoria e suporte.
7. Guardar a chave privada num serviço/cofre separado e com registo de emissão.

## Integração recomendada

1. Manter o trial atual enquanto se fecha a política comercial.
2. Criar um `LicenseService` na camada da aplicação que obtenha hora confiável,
   fingerprint e token através de interfaces explícitas.
3. Validar a licença uma vez no arranque e renovar o estado em segundo plano.
4. Aplicar permissões por módulo sem esconder ou apagar dados do cliente.
5. Permitir sempre acesso a backup, exportação e contacto de suporte quando uma
   licença expira.
6. Testar relógio adulterado, token corrompido, troca de equipamento, ausência de
   rede e atualização de versão antes de ativar bloqueios comerciais.

## Regras de segurança

- Nunca guardar a chave privada no Git ou no executável.
- Nunca usar apenas um JSON editável como prova de licença.
- Nunca depender apenas do relógio local.
- Nunca associar licenciamento às credenciais MySQL do cliente.
- Nunca apagar dados por expiração de licença.
- O estado da licença não substitui autenticação nem permissões de utilizador.
