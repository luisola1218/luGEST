# IA focada em Produtos e Matéria-Prima

O LuGEST funciona integralmente sem IA. A inteligência opcional está concentrada
nos dois fluxos em que acrescenta valor verificável:

- identificação, normalização e classificação de Produtos;
- preparação de novos lotes de Matéria-Prima a partir de linguagem natural.

Nenhuma proposta é gravada automaticamente. O utilizador revê os campos e
confirma a ficha antes de alterar a base de dados.

## Opção recomendada para clientes

Em produção, os postos devem chamar um gateway HTTPS controlado pela LuGEST:

```ini
LUGEST_AI_ENDPOINT=https://ia.exemplo.pt/structured
LUGEST_AI_ACCESS_TOKEN=token-individual-da-instalacao
LUGEST_AI_TIMEOUT_SECONDS=45
```

A chave do fornecedor de IA permanece apenas no gateway. Nunca deve ser
incluída no executável, no pacote comercial ou no `lugest.env` entregue ao
cliente.

## OpenAI num ambiente interno controlado

Para desenvolvimento ou num servidor pertencente ao cliente:

```ini
LUGEST_OPENAI_API_KEY=chave-do-projeto-openai
LUGEST_OPENAI_MODEL=gpt-5.6-sol
LUGEST_OPENAI_BASE_URL=https://api.openai.com/v1
LUGEST_AI_TIMEOUT_SECONDS=45
```

A subscrição ChatGPT não é uma credencial da API. É necessária uma chave de
projeto da plataforma OpenAI ou o gateway seguro descrito acima.

## Alternativa local gratuita com Ollama

1. Instala o Ollama para Windows.
2. Executa:

   ```powershell
   ollama pull qwen3:4b
   ```

3. Configura:

   ```ini
   LUGEST_OLLAMA_URL=http://127.0.0.1:11434
   LUGEST_OLLAMA_MODEL=qwen3:4b
   LUGEST_AI_TIMEOUT_SECONDS=45
   ```

O Ollama é usado como alternativa quando o serviço cloud não está configurado
ou não responde. O primeiro pedido pode demorar enquanto o modelo é carregado.
Não exponhas a porta 11434 diretamente à Internet.

## Produtos

Em **Produtos**, escreve a descrição e seleciona **Pesquisar com IA**. O motor
propõe descrição normalizada, categoria, subcategoria, tipo, fabricante, modelo,
dimensões e atributos industriais. A resposta segue um esquema fechado e é
validada pelo LuGEST antes de preencher a ficha.

## Matéria-Prima

No campo de criação assistida podes escrever, por exemplo:

```text
Cria stock de chapa S235JR 15 mm, formato 3000x1500, 10 unidades, lote externo 9288X20029
```

O LuGEST prepara a ficha e aplica também as suas regras locais de materiais,
perfis normalizados, dimensões e pesos. O lote só é criado depois de
**Criar lote confirmado**.

## Prioridade e funcionamento sem IA

Ordem dos motores:

1. gateway HTTPS LuGEST;
2. OpenAI configurada no servidor controlado;
3. Google Gemini, apenas quando explicitamente configurada;
4. Ollama local ou na rede;
5. regras determinísticas do LuGEST.

Sem qualquer motor, os formulários, validações, cálculos e restantes módulos
continuam operacionais; apenas não aparece uma proposta externa avançada.
