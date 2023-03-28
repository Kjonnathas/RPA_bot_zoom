O projeto foi desenvolvido todo utilizando apenas o Selenium como ferramenta para manipulação do navegador e raspagem de dados.

Lógica do script:

- Abre o navegador do Chrome;

- Entra no site do Zoom;

- Procura pelo produto desejado;

- Varre todas as páginas encontradas e coleta as informações de nome do produto, preço, parcelamento, loja anunciante e link do anúncio;

- Cria dois arquivos em .xlsx, um para armazenar todos os produtos encontrados e outro para armazenar apenas os produtos encontrados dentro da faixa de preço estabelecida;

- Caso encontre algum produto dentro da faixa de preço estabelecida, dispara um e-mail com as informações no corpo do e-mail (se for apenas um único produto encontrado) ou anexa um arquivo em .xlsx (se for mais de um produto encontrado) e aciona a notificação do Windows para alertar;

- Fecha o navegador e finaliza o código.

Foi criado um ambiente virtual para desenvolver o código e para instalar apenas as bibliotecas utilizadas no projeto. Além disso, foi criado um executável para que qualquer pessoa possa utilizar, indenpendente de qual máquina esteja rodando.

Passo a passo de como utilizar o código:

1. Baixe a pasta "dist";

2. Siga o caminho: dist >> Bot_zoom e, posteriormente, procure pelo arquivo "Produto.txt";

3. Abra o arquivo "Produto.txt" e insira o nome do produto desejado, preco mínimo e preço máximo. Ex.: Televisão Samsung 50 polegadas, 2000, 3500

Obs.: O arquivo "Produto.txt" precisa estar justamente como demonstrado no exemplo acima, senão irá ocorrer erro na manipulação dos dados.