Конвертация xml в excel через mvn xml:transform

Путь к папке входной xml указывается в maven проперти ${xmlDir}

Путь к xsl указывается в maven проперти ${stylesheetPath}

После mvn xml:transform в ${basedir}\target\generated-resources\xml\xslt будет создан index.xls

Attention: если в ${xmlDir} много xml, то в index.xls будут данные последней xml