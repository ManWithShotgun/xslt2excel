<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="2.0"
                xmlns="urn:schemas-microsoft-com:office:spreadsheet"
                xmlns:msxsl="urn:schemas-microsoft-com:xslt"
                xmlns:x="urn:schemas-microsoft-com:office:excel"
                xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
    <xsl:output method="xml" indent="yes"/>
    <xsl:template match="/">
        <xsl:processing-instruction name="mso-application">progid="Excel.Sheet"</xsl:processing-instruction>
        <Workbook>
            <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
                <WindowHeight>10005</WindowHeight>
                <WindowWidth>10005</WindowWidth>
                <WindowTopX>120</WindowTopX>
                <WindowTopY>135</WindowTopY>
                <ProtectStructure>False</ProtectStructure>
                <ProtectWindows>False</ProtectWindows>
            </ExcelWorkbook>
            <Styles>
                <Style ss:ID="head_cell_style">
                    <Font ss:Bold="1" x:Family="Swiss"/>
                    <Borders>
                        <Border ss:LineStyle="Continuous" ss:Position="Left" ss:Weight="1"/>
                        <Border ss:LineStyle="Continuous" ss:Position="Top" ss:Weight="1"/>
                    </Borders>
                    <Interior ss:Color="#C0C0C0" ss:Pattern="Solid"/>
                    <Alignment ss:Vertical="Bottom" ss:WrapText="1"/>
                </Style>
                <Style ss:ID="solid_row_style">
                    <Interior ss:Color="#C0C0C0" ss:Pattern="Solid"/>
                    <Borders>
                        <Border ss:LineStyle="Continuous" ss:Position="Left" ss:Weight="1"/>
                        <Border ss:LineStyle="Continuous" ss:Position="Right" ss:Weight="1"/>
                    </Borders>
                </Style>
                <Style ss:ID="solid_row_top_style">
                    <Interior ss:Color="#C0C0C0" ss:Pattern="Solid"/>
                    <Borders>
                        <Border ss:LineStyle="Continuous" ss:Position="Left" ss:Weight="1"/>
                        <Border ss:LineStyle="Continuous" ss:Position="Right" ss:Weight="1"/>
                        <Border ss:LineStyle="Continuous" ss:Position="Top" ss:Weight="1"/>
                    </Borders>
                </Style>
                <Style ss:ID="value_row_style">
                    <Borders>
                        <Border ss:LineStyle="Continuous" ss:Position="Left" ss:Weight="1"/>
                        <Border ss:LineStyle="Continuous" ss:Position="Right" ss:Weight="1"/>
                        <Border ss:LineStyle="Continuous" ss:Position="Top" ss:Weight="1"/>
                        <Border ss:LineStyle="Continuous" ss:Position="Bottom" ss:Weight="1"/>
                    </Borders>
                </Style>
            </Styles>
            <xsl:apply-templates select="items"/>
        </Workbook>
    </xsl:template>
    <xsl:template match="items">
        <Worksheet ss:Name="{@type}">
            <Table ss:DefaultRowHeight="14" ss:ExpandedColumnCount="6" x:FullColumns="1" x:FullRows="1">
                <Column ss:Index="1" ss:Width="100"/>
                <Column ss:Index="2" ss:Width="300"/>
                <Column ss:Index="3" ss:Width="100"/>
                <!--Header Row-->
                <Row>
                    <Cell ss:StyleID="head_cell_style">
                        <Data ss:Type="String">
                            user
                        </Data>
                    </Cell>
                    <Cell ss:StyleID="head_cell_style">
                        <Data ss:Type="String">
                            Task description
                        </Data>
                    </Cell>
                    <Cell ss:StyleID="head_cell_style">
                        <Data ss:Type="String">
                            Time spent (h)
                        </Data>
                    </Cell>
                </Row>
                <!--Content-->
                <xsl:for-each-group select="item" group-by="field[@name='userName']">

                    <xsl:for-each-group select="current-group()" group-by="field[@name='description']">

                        <Row>
                            <Cell  ss:StyleID="value_row_style">
                                <Data ss:Type="String">
                                    <xsl:value-of select="field[@name='userName']"/>
                                </Data>
                            </Cell>
                            <Cell  ss:StyleID="value_row_style">
                                <Data ss:Type="String">
                                    <xsl:value-of select="field[@name='description']"/>
                                </Data>
                            </Cell>
                            <Cell  ss:StyleID="value_row_style">
                                <Data ss:Type="Number">
                                    <xsl:value-of select="sum(current-group()/field[@name='activityHours'])"/>
                                </Data>
                            </Cell>
                        </Row>

                    </xsl:for-each-group>

                </xsl:for-each-group>

            </Table>
        </Worksheet>
    </xsl:template>

</xsl:stylesheet>