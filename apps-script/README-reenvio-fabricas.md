# Reenvio: planilhas das fábricas

Script: `reenvio-fabricas.gs`. Ele é separado do script da aba **Reenvios** e não mexe nele.

O que ele faz:

- **Consulta.** O site (aba Reenvio, cartão "Acompanhar reenvio") busca pelo order number e mostra a fábrica, o status, o tracking, a data de envio, o produto, o protocolo e o ticket.
  - E-mail e endereço do cliente **nunca** saem da planilha por aqui.
- **Registro.** Depois que um reenvio é gravado na aba Reenvios, o site manda a mesma linha para a planilha da fábrica certa, com status "Sent to factory" e a data do dia.
  - Fullstack ainda não tem planilha: o site só avisa.

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
