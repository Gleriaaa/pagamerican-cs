# Reenvio: planilhas das fábricas

Script: `reenvio-fabricas.gs`. Ele é separado do script da aba **Reenvios** e não mexe nele.

O que ele faz:

- **Consulta.** O site (aba Reenvio, cartão "Acompanhar reenvio") busca pelo order number e mostra a fábrica, o status, o tracking, a data de envio, o produto, o protocolo e o ticket.
  - E-mail e endereço do cliente **nunca** saem da planilha por aqui.
- **Validação.** Depois que um reenvio é gravado na aba Reenvios, o site manda a linha para a planilha interna **Validação de reenvios**, com "Aguardando validação", e a validadora recebe um e-mail. As fábricas não têm acesso a essa planilha.
- **Registro na fábrica.** Só quando a validadora escolhe **Aprovado** a linha vai para a planilha da fábrica certa, com status "Sent to factory" e a data do dia, e a fábrica recebe o aviso.
  - Fullstack ainda não tem planilha: a linha fica como "Aprovado: enviar manualmente".

## Colunas esperadas (em qualquer ordem, nas 5 primeiras linhas)

Order Number · Product · Quantity · Shipping Address · Customer Email · Customer Name · Account / Vendor · Protocol · Ticket · Date sent to factory · Reshipment tracking · Status · Notes

Os nomes em português (Produto, Protocolo, Data do envio à fábrica, Tracking do reenvio, Observações…) também funcionam.

A fábrica preenche **Reshipment tracking** e **Status**. Quando o status contém "Delivered" (ou "Entregue"), o site mostra o selo em verde. Quando contém "Problem", "Returned", "Lost", "Cancelled" ou os equivalentes em português, mostra em vermelho.

## Como publicar (uma vez)

1. Suba os 3 arquivos no Google Drive, abra cada um e use **Arquivo → Salvar como Planilhas Google**. O script não abre arquivo .xlsx: o ID que vale é o da versão Planilhas Google.
2. Copie o ID de cada planilha: é o trecho do link entre `/d/` e `/edit`.
3. Em <https://script.google.com>, crie um **Novo projeto**, apague o conteúdo e cole o `reenvio-fabricas.gs`.
4. Troque os 3 `COLE_O_ID_DA_PLANILHA_...` pelos IDs. Se os dados não estiverem na primeira aba, preencha `aba` com o nome dela.
5. Clique em **Implantar → Nova implantação → App da Web**:
   - Executar como: **eu**.
   - Quem pode acessar: **Qualquer pessoa**.
   - Autorize o acesso às planilhas.
6. Copie a URL que termina em `/exec` e cole em `RS_TRK_URL` no `index.html`.

A chave (`CHAVE` no script, `RS_TRK_KEY` no site) já vem igual nos dois arquivos. Não troque só de um lado.

Ao alterar o script depois, use **Implantar → Gerenciar implantações → editar → Nova versão**. Assim a URL continua a mesma.

## Compartilhar com as fábricas

Cada fábrica recebe acesso de **editor só à própria planilha**. O script roda com a sua conta, então as fábricas não precisam ter acesso a ele.

## Avisos por e-mail para as fábricas

Os e-mails vão em inglês para **todos os editores da planilha da fábrica**, menos o dono. Compartilhou como Editor, a pessoa passa a receber. Para parar, tire o acesso dela. Os e-mails nunca trazem endereço nem e-mail do cliente.

- **Na hora:** cada reenvio novo enviado pelo site gera um e-mail "New reshipment request – Order ..." com pedido, produto, quantidade, protocolo e o link da planilha.
- **Todo dia às 9h (horário de Nova York):** um lembrete com os reenvios sem tracking há 2 dias ou mais, que não estejam entregues nem cancelados.
  - Para ligar o lembrete, rode **uma vez** a função `instalarLembrete`: no editor, escolha a função na lista ao lado de "Depurar" e clique em **Executar**.
  - Para mudar a hora, os dias ou o fuso, ajuste `LEMBRETE_HORA`, `LEMBRETE_DIAS_MIN` e `LEMBRETE_FUSO` e rode `instalarLembrete` de novo.
- **Cópia para a PagAmerican:** coloque os e-mails em `AVISO_COPIA`.

## Validação antes da fábrica

Todo reenvio passa por uma validação antes de chegar à fábrica. A validação fica numa **planilha separada**, porque o Google não deixa esconder uma aba de quem tem acesso à planilha: se ficasse numa aba da planilha da fábrica, a fábrica veria.

### Como funciona

1. O agente envia pelo site, como sempre. A linha entra na aba Reenvios e na planilha **Validação de reenvios — PagAmerican (interno)**, com a Situação "Aguardando validação".
2. A validadora recebe um e-mail "Reenvio para validar – Order …" com o motivo, o agente, o ticket e quantos reenvios essa order já teve, e um link direto para a linha. O e-mail não traz endereço nem e-mail do cliente.
3. Na planilha, ela escolhe na coluna **Decisão**:
   - **Aprovado:** na hora, a linha vai para a planilha da fábrica e a fábrica recebe o aviso. A Situação muda para "Enviado à … em …".
   - **Reprovado:** ela escreve o motivo em **Motivo da reprovação**. Nada vai para a fábrica.
4. No site, o "Acompanhar reenvio" mostra "Aguardando validação", "Reprovado na validação: motivo" ou, depois de aprovado, o tracking da fábrica.

Só as colunas **Decisão** e **Motivo da reprovação** podem ser editadas; o resto é protegido. Depois que um reenvio foi enviado à fábrica, a decisão não volta atrás pela planilha: para cancelar, fale com a fábrica.

A coluna **Reenvios anteriores** mostra "⚠" quando a order já tem reenvio numa planilha de fábrica ou na validação.

Quem recebe o e-mail: **todos os editores da planilha de validação, menos o dono**. Para incluir alguém, compartilhe como Editor; para tirar, remova o acesso. Cópia fixa: `VALIDACAO_COPIA`.

### Como ligar (uma vez)

1. No projeto do script, cole o `reenvio-fabricas.gs` novo e salve.
2. Escolha a função **configurarValidacao** na lista ao lado de "Depurar" e clique em **Executar**. Autorize os acessos que o Google pedir.
   - Ela cria a planilha, protege as colunas, coloca a lista Aprovado/Reprovado e instala o gatilho de edição e a varredura a cada 10 minutos.
   - O link da planilha aparece no **Registro de execução**.
3. Abra a planilha e compartilhe com a validadora como **Editor** (por e-mail, nunca por link). Nunca compartilhe com as fábricas.
4. Publique em **Implantar → Gerenciar implantações → lápis → Nova versão**. A URL continua a mesma.
5. Suba o `index.html` novo no site.

Teste: faça um reenvio de teste pelo site, confira o e-mail da validadora, aprove e veja a linha chegar na planilha da fábrica. Depois apague a linha de teste nas duas planilhas.

A partir da ativação, nenhum reenvio vai direto para a fábrica. Para voltar ao fluxo antigo, rode **desligarValidacao** (a planilha não é apagada).
