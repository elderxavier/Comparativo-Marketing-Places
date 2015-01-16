GERAR RELATORIO DE ESTOQUE EXTRA.COM

1 - consultar o Banco de dados do Magento com a instrução:
	SELECT `catalog_product_entity`.`sku`, `cataloginventory_stock_item`.`qty`, `cataloginventory_stock_item`.`is_in_stock` 
	FROM `catalog_product_entity` 
	INNER JOIN `cataloginventory_stock_item` ON  `catalog_product_entity`.`entity_id` =`cataloginventory_stock_item`.`product_id` 
	WHERE `catalog_product_entity`.`sku`<> "NULL"
	ORDER BY `catalog_product_entity`.`sku` ASC

2 - Exporte uma consulta em formado CSV for MS Excel

3 - Copie a a tabela da consulta do Banco de Dados para Planilha 2(Magento Estoque) e insira os titulos: sku = A; qty=B; is_in_stock= C;  
	Obs: (A Planilha 2 pode ser renomeado mais não pode alterar a ordem, ele deve ter indice 2)

4 - Copie a tabela de estoque do painel Admin do Extra para a Planilha 3 (Extra Estoque)

5 - Acione o botão  "Ajusta tabela  para Tabela Extra" para ajustar o partner

6 - Escolha a opção de consulta

8 - Aguarde a mesnagem de finalização 
	obs: Devido o tamanho da pesquisa o tempo de execução pode ser superior a 5 min e Excel irá parar de responder durante este tempo, depois retornará a funcionar

	
GERAR RELATORIO DE PREÇOS EXTRA.COM
	
	
1 - consultar o Banco de dados do Magento com a instrução:
	SELECT `catalog_product_entity`.`sku`,`catalog_product_entity_decimal`.`value`, price.value AS price
	FROM catalog_product_entity ,`catalog_product_entity_decimal`
	LEFT JOIN catalog_product_entity_decimal price ON (price.`entity_type_id`=4) AND (price.`attribute_id`=64)
	WHERE `catalog_product_entity`.`entity_id`=`catalog_product_entity_decimal`.`entity_id` AND `catalog_product_entity_decimal`.`attribute_id`=75 
	AND `catalog_product_entity`.`sku`<> "NULL"
	ORDER BY `catalog_product_entity`.`sku` ASC

2 - Exporte uma consulta em formado CSV for MS Excel

3 - Copie a a tabela da consulta do Banco de Dados para Planilha 4(Magento Preco) e insira os titulos: sku = A; value = B; price = C;  
	Obs: (A Planilha 4 pode ser renomeado mais não pode alterar a ordem, ele deve ter indice 2)

4 - Copie a tabela de estoque do painel Admin do Extra para a planilha 5 (Extra Preco)

5 - Acione o botão  "Ajusta tabela  para Tabela Extra" para ajustar o partner

6 - Escolha a opção de consulta

8 - Aguarde a mesnagem de finalização 
	obs: Devido o tamanho da pesquisa o tempo de execução pode ser superior a 5 min e Excel irá parar de responder durante este tempo, depois retornará a funcionar

	