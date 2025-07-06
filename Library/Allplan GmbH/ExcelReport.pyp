<?xml version="1.0" encoding="utf-8"?>
<Element>
    <Script>
        <Name>allplan_gmbh\ExcelReport.py</Name>
        <Title>ExcelReport</Title>
        <Version>1.0</Version>
        <Interactor>True</Interactor>
        <ReadLastInput>False</ReadLastInput>
    </Script>
    <Page>
        <Name>Page1</Name>
        <Text>CreateExcelReport</Text>
                <Parameter>
                    <Name>Image</Name>
                    <Value>ExcelReport.svg</Value>
                    <Text>Grafik</Text>
                    <Orientation>Middle</Orientation>
                    <ValueType>Picture</ValueType>
                    <Visible>True</Visible>
                </Parameter>
                <Parameter>
                    <Name>Separator</Name>
                    <ValueType>Separator</ValueType>
                    <Visible>True</Visible>
                </Parameter>

<!-- selection of the report kind -->
            <Parameter>
                <Name>Report_kind</Name>
                <TextId>1001</TextId>
                <Text>Kind of Report</Text>
                <Value>param</Value>
                <ValueType>RadioButtonGroup</ValueType>
                <Parameter>
                    <Name>param_report</Name>
                    <TextId>1002</TextId>
                    <Text>Objects parameter report</Text>
                    <Value>param</Value>
                    <ValueType>RadioButton</ValueType>
                </Parameter>
                <Parameter>
                    <Name>qto_report</Name>
                    <TextId>1003</TextId>
                    <Text>Objects quantity report (QTO)</Text>
                    <Value>qto</Value>
                    <ValueType>RadioButton</ValueType>
                </Parameter>
            </Parameter>
            <Parameter>
                <Name>Expander</Name>
                <TextId>1004</TextId>
                <Text>Exel file</Text>
                <ValueType>Expander</ValueType>
                <Value>False</Value>
<!-- options for the Excel table -->
                <Parameter>
                    <Name>Row10</Name>
                    <TextId>1005</TextId>
                    <Text>Excel tabel</Text>
                    <ValueType>Row</ValueType>
                    <Parameter>
                        <Name>tabel_handling</Name>
                        <TextId>1005</TextId>
                        <Text>Excel tabel</Text>
                        <Value>create</Value>
                        <ValueType>RadioButtonGroup</ValueType>
                        <Parameter>
                            <Name>create_new</Name>
                            <TextId>1006</TextId>
                            <Text>new</Text>
                            <Value>create</Value>
                            <ValueType>RadioButton</ValueType>
                        </Parameter>
                        <Parameter>
                            <Name>extend_existing</Name>
                            <TextId>1007</TextId>
                            <Text>open</Text>
                            <Value>use</Value>
                            <ValueType>RadioButton</ValueType>
                        </Parameter>
                    </Parameter>
                </Parameter>
<!-- selection of the Excel tabel -->
                <Parameter>
                    <Name>Row1</Name>
                    <TextId>1008</TextId>
                    <Text>Select table</Text>
                    <ValueType>Row</ValueType>
                    <Parameter>
                        <Name>read_tabel_path</Name>
                        <TextId>1008</TextId>
                        <Text>Select table</Text>
                        <Value></Value>
                        <ValueType>String</ValueType>
                        <ValueDialog>OpenFileDialog</ValueDialog>
                        <FileFilter>Excel file(*.xlsx)|*.xlsx|</FileFilter>
                        <FileExtension>xlsx</FileExtension>
                        <Visible>tabel_handling == "use"</Visible>
                        <DefaultDirectories>std|prj</DefaultDirectories>
                    </Parameter>
                </Parameter>
                <Parameter>
                    <Name>save_tabel_path</Name>
                    <TextId>1009</TextId>
                    <Text>Select path</Text>
                    <Value></Value>
                    <ValueType>String</ValueType>
                    <ValueDialog>SaveFileDialog</ValueDialog>
                    <FileFilter>Excel file(*.xlsx)|*.xlsx|</FileFilter>
                    <FileExtension>xlsx</FileExtension>
                    <Visible>tabel_handling == "create"</Visible>
                    <DefaultDirectories>std|prj</DefaultDirectories>
                </Parameter>
                <Parameter>
                    <Name>Row2</Name>
                    <TextId>1010</TextId>
                    <Text>Table sheet</Text>
                    <ValueType>Row</ValueType>
                    <Parameter>
                        <Name>read_sheet_name</Name>
                        <Text>Sheet</Text>
                        <Value></Value>
                        <ValueList></ValueList>
                        <ValueType>StringComboBox</ValueType>
                        <Visible>tabel_handling == "use"</Visible>
                    </Parameter>
                </Parameter>
                <Parameter>
                    <Name>Separator</Name>
                    <ValueType>Separator</ValueType>
                    <Visible>False</Visible>
                </Parameter>
                <Parameter>
                    <Name>Row3</Name>
                    <TextId>1011</TextId>
                    <Text>Line to start</Text>
                    <ValueType>Row</ValueType>
                    <Parameter>
                        <Name>start_row</Name>
                        <TextId>1012</TextId>
                        <Text>first line</Text>
                        <Value>2</Value>
                        <ValueType>Integer</ValueType>
                    </Parameter>
                </Parameter>
                <Parameter>
                    <Name>Row4</Name>
                    <TextId>1013</TextId>
                    <Text>Column to start</Text>
                    <ValueType>Row</ValueType>
                    <Parameter>
                        <Name>start_column</Name>
                        <TextId>1014</TextId>
                        <Text>first column</Text>
                        <Value>1</Value>
                        <ValueType>Integer</ValueType>
                    </Parameter>
                </Parameter>
            </Parameter>
<!-- Festlegung Attribute für Parameter report -->
            <Parameter>
                <Name>Expander</Name>
                <Text>Parameter</Text>
                <ValueType>Expander</ValueType>
                <Value>False</Value>
                <Visible>Report_kind =="param"</Visible>
                <Parameter>
                    <Name>Row4</Name>
                    <TextId>1015</TextId>
                    <Text>Object identifier</Text>
                    <ValueType>Row</ValueType>
                    <Parameter>
                        <Name>object_ident</Name>
                        <TextId>1016</TextId>
                        <Text>Identifier attribute</Text>
                        <Value> </Value>
                        <ValueList> |Allright_ID|IFC_ID|Choose...</ValueList>
                        <ValueList_TextIds> |1017|1018|1019</ValueList_TextIds>
                        <ValueType>StringComboBox</ValueType>
                    </Parameter>
                </Parameter>
                <Parameter>
                    <Name>Row5</Name>
                    <Text>Kennerattribut</Text>
                    <ValueType>Row</ValueType>
                    <Parameter>
                        <Name>object_ident_ID</Name>
                        <Text></Text>
                        <Value></Value>
                        <ValueType>AttributeId</ValueType>
                        <ValueDialog>AttributeSelectionInsert</ValueDialog>
                        <Visible>object_ident == "Choose..." or object_ident == "Wählen..."</Visible>
                    </Parameter>
                </Parameter>
                <Parameter>
                    <Name>Separator</Name>
                    <ValueType>Separator</ValueType>
                </Parameter>
                <Parameter>
                    <Name>Report_attribs</Name>
                    <Text>Reportparameter</Text>
                    <ValueType>Text</ValueType>
                </Parameter>
                <Parameter>
                    <Name>AttributeIDList</Name>
                    <Text></Text>
                    <Value>[0]</Value>
                    <ValueType>AttributeId</ValueType>
                    <ValueDialog>AttributeSelection</ValueDialog>
                    <ValueListStartRow>1</ValueListStartRow>
                </Parameter>
                <Parameter>
                    <Name>Separator</Name>
                    <ValueType>Separator</ValueType>
                </Parameter>
                <Parameter>
                    <Name>Row10</Name>
                    <Text>Gruppierung nach</Text>
                    <ValueType>Row</ValueType>
                    <Parameter>
                        <Name>group_attrib</Name>
                        <Text>Gruppierung</Text>
                        <Value></Value>
                        <ValueList></ValueList>
                        <ValueType>StringComboBox</ValueType>
                    </Parameter>
                </Parameter>
                <Parameter>
                    <Name>AttributeInfoList</Name>
                    <Text>Attribute IDs</Text>
                    <Value>[]</Value>
                    <ValueType>Integer</ValueType>
                    <Enable>True</Enable>
                    <ValueListStartRow>1</ValueListStartRow>
                    <Visible>len(AttributeIDList) &gt; 1</Visible>
                </Parameter>
            </Parameter>
    <!-- Festlegung Parameter für Mengen report -->
            <Parameter>
                <Name>Expander</Name>
                <Text>QTO Parameter</Text>
                <ValueType>Expander</ValueType>
                <Value>False</Value>
                <Visible>Report_kind =="qto"</Visible>
                <Parameter>
                    <Name>Row15</Name>
                    <Text>QTO Attribut</Text>
                    <ValueType>Row</ValueType>
                    <Parameter>
                        <Name>qto_ident_attrib</Name>
                        <Text></Text>
                        <Value>508</Value>
                        <ValueType>AttributeId</ValueType>
                        <ValueDialog>AttributeSelectionInsert</ValueDialog>
                        <Visible></Visible>
                    </Parameter>
                </Parameter>
                <Parameter>
                    <Name>Sep</Name>
                    <Text>Mengenparameter</Text>
                    <ValueType>Separator</ValueType>
                </Parameter>
                <Parameter>
                    <Name>QTO_attribs</Name>
                    <Text>Mengenparameter</Text>
                    <ValueType>Text</ValueType>
                </Parameter>
                <Parameter>
                    <Name>QTOAttributeList</Name>
                    <Text></Text>
                    <Value>[0]</Value>
                    <ValueType>AttributeId</ValueType>
                    <ValueDialog>AttributeSelection</ValueDialog>
                    <ValueListStartRow>1</ValueListStartRow>
                </Parameter>
            </Parameter>
<!-- Auswahl Objekte -->
            <Parameter>
                <Name>Expander</Name>
                <Text>Elementauswahl</Text>
                <ValueType>Expander</ValueType>
                <Value>False</Value>
                <Parameter>
                    <Name>Row2</Name>
                    <Text>Objekte</Text>
                    <ValueType>Row</ValueType>
                    <Parameter>
                        <Name>element_selection</Name>
                        <Text>auswählen</Text>
                        <EventId>1002</EventId>
                        <ValueType>Button</ValueType>
                    </Parameter>
                </Parameter>
            </Parameter>
<!-- Erstellung der Tabelle -->
            <Parameter>
                <Name>Expander</Name>
                <Text>Erstellung</Text>
                <ValueType>Expander</ValueType>
                <Value>False</Value>
                <Parameter>
                    <Name>Row48</Name>
                    <Text>Excel Report</Text>
                    <ValueType>Row</ValueType>
                    <Parameter>
                        <Name>create_excel_report</Name>
                        <Text>erstellen</Text>
                        <EventId>1005</EventId>
                        <ValueType>Button</ValueType>
                    </Parameter>

                </Parameter>

            </Parameter>
    </Page>
</Element>
