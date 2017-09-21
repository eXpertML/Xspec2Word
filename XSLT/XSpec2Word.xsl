<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:math="http://www.w3.org/2005/xpath-functions/math"
  xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl"
  xmlns:x="http://www.jenitennison.com/xslt/xspec"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:w2006="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:wx="http://schemas.microsoft.com/office/word/2003/auxHint"
  xmlns:x2w="http://ns.expertml.com/xslt/x2w"
  exclude-result-prefixes="#all"
  expand-text="yes"
  version="3.0">
  
  <!--<xd:doc scope="stylesheet">
    <xd:desc>
      <xd:p><xd:b>Created on:</xd:b> Feb 4, 2017</xd:p>
      <xd:p><xd:b>Author:</xd:b> TFJH</xd:p>
      <xd:p>This stylesheet creates a (2003) Microsoft Word™ document from an XSpec test; any XML that is compatible with word is output in native Word format, the formatting for which can be supplied separately.</xd:p>
    </xd:desc>
  </xd:doc>-->

  <xsl:output indent="yes"/>
  
  <xsl:preserve-space elements="w:t"/>
  
  <xsl:param name="fonts" as="element(w:fonts)">
    <w:fonts>
      <w:defaultFonts w:ascii="Calibri" w:fareast="Calibri" w:h-ansi="Calibri" w:cs="Times New Roman"/>
      <w:font w:name="Times New Roman">
        <w:panose-1 w:val="02020603050405020304"/>
        <w:charset w:val="00"/>
        <w:family w:val="auto"/>
        <w:pitch w:val="variable"/>
        <w:sig w:usb-0="E0002AEF" w:usb-1="C0007841" w:usb-2="00000009" w:usb-3="00000000" w:csb-0="000001FF" w:csb-1="00000000"/>
      </w:font>
      <w:font w:name="Arial">
        <w:panose-1 w:val="020B0604020202020204"/>
        <w:charset w:val="00"/>
        <w:family w:val="auto"/>
        <w:pitch w:val="variable"/>
        <w:sig w:usb-0="E0002AFF" w:usb-1="C0007843" w:usb-2="00000009" w:usb-3="00000000" w:csb-0="000001FF" w:csb-1="00000000"/>
      </w:font>
      <w:font w:name="Courier New">
        <w:panose-1 w:val="02070309020205020404"/>
        <w:charset w:val="00"/>
        <w:family w:val="auto"/>
        <w:pitch w:val="variable"/>
        <w:sig w:usb-0="E0002AFF" w:usb-1="C0007843" w:usb-2="00000009" w:usb-3="00000000" w:csb-0="000001FF" w:csb-1="00000000"/>
      </w:font>
      <w:font w:name="Symbol">
        <w:panose-1 w:val="05050102010706020507"/>
        <w:charset w:val="02"/>
        <w:family w:val="auto"/>
        <w:pitch w:val="variable"/>
        <w:sig w:usb-0="00000000" w:usb-1="10000000" w:usb-2="00000000" w:usb-3="00000000" w:csb-0="80000000" w:csb-1="00000000"/>
      </w:font>
      <w:font w:name="Helvetica">
        <w:panose-1 w:val="00000000000000000000"/>
        <w:charset w:val="00"/>
        <w:family w:val="auto"/>
        <w:pitch w:val="variable"/>
        <w:sig w:usb-0="E00002FF" w:usb-1="5000785B" w:usb-2="00000000" w:usb-3="00000000" w:csb-0="0000019F" w:csb-1="00000000"/>
      </w:font>
      <w:font w:name="Wingdings">
        <w:panose-1 w:val="05000000000000000000"/>
        <w:charset w:val="02"/>
        <w:family w:val="auto"/>
        <w:pitch w:val="variable"/>
        <w:sig w:usb-0="00000000" w:usb-1="10000000" w:usb-2="00000000" w:usb-3="00000000" w:csb-0="80000000" w:csb-1="00000000"/>
      </w:font>
      <w:font w:name="Cambria Math">
        <w:panose-1 w:val="02040503050406030204"/>
        <w:charset w:val="00"/>
        <w:family w:val="auto"/>
        <w:pitch w:val="variable"/>
        <w:sig w:usb-0="E00002FF" w:usb-1="420024FF" w:usb-2="00000000" w:usb-3="00000000" w:csb-0="0000019F" w:csb-1="00000000"/>
      </w:font>
      <w:font w:name="Calibri">
        <w:panose-1 w:val="020F0502020204030204"/>
        <w:charset w:val="00"/>
        <w:family w:val="auto"/>
        <w:pitch w:val="variable"/>
        <w:sig w:usb-0="E00002FF" w:usb-1="4000ACFF" w:usb-2="00000001" w:usb-3="00000000" w:csb-0="0000019F" w:csb-1="00000000"/>
      </w:font>
      <w:font w:name="Cambria">
        <w:panose-1 w:val="02040503050406030204"/>
        <w:charset w:val="00"/>
        <w:family w:val="auto"/>
        <w:pitch w:val="variable"/>
        <w:sig w:usb-0="E00002FF" w:usb-1="400004FF" w:usb-2="00000000" w:usb-3="00000000" w:csb-0="0000019F" w:csb-1="00000000"/>
      </w:font>
      <w:font w:name="MS Mincho">
        <w:panose-1 w:val="02020609040205080304"/>
        <w:charset w:val="80"/>
        <w:family w:val="auto"/>
        <w:pitch w:val="variable"/>
        <w:sig w:usb-0="E00002FF" w:usb-1="6AC7FDFB" w:usb-2="08000012" w:usb-3="00000000" w:csb-0="0002009F" w:csb-1="00000000"/>
      </w:font>
      <w:font w:name="QuaySansITCStd-MediumIt">
        <w:altName w:val="Times New Roman"/>
        <w:panose-1 w:val="00000000000000000000"/>
        <w:charset w:val="4D"/>
        <w:family w:val="auto"/>
        <w:notTrueType/>
        <w:pitch w:val="default"/>
        <w:sig w:usb-0="00000003" w:usb-1="00000000" w:usb-2="00000000" w:usb-3="00000000" w:csb-0="00000001" w:csb-1="00000000"/>
      </w:font>
      <w:font w:name="QuaySansITCStd-Black">
        <w:altName w:val="Times New Roman"/>
        <w:panose-1 w:val="00000000000000000000"/>
        <w:charset w:val="4D"/>
        <w:family w:val="auto"/>
        <w:notTrueType/>
        <w:pitch w:val="default"/>
        <w:sig w:usb-0="00000003" w:usb-1="00000000" w:usb-2="00000000" w:usb-3="00000000" w:csb-0="00000001" w:csb-1="00000000"/>
      </w:font>
      <w:font w:name="QuaySansITCStd-Book">
        <w:altName w:val="Times New Roman"/>
        <w:panose-1 w:val="00000000000000000000"/>
        <w:charset w:val="4D"/>
        <w:family w:val="auto"/>
        <w:notTrueType/>
        <w:pitch w:val="default"/>
        <w:sig w:usb-0="00000003" w:usb-1="00000000" w:usb-2="00000000" w:usb-3="00000000" w:csb-0="00000001" w:csb-1="00000000"/>
      </w:font>
      <w:font w:name="Tahoma">
        <w:panose-1 w:val="020B0604030504040204"/>
        <w:charset w:val="00"/>
        <w:family w:val="auto"/>
        <w:pitch w:val="variable"/>
        <w:sig w:usb-0="E1002EFF" w:usb-1="C000605B" w:usb-2="00000029" w:usb-3="00000000" w:csb-0="000101FF" w:csb-1="00000000"/>
      </w:font>
      <w:font w:name="Myriad Pro">
        <w:altName w:val="Arial"/>
        <w:panose-1 w:val="00000000000000000000"/>
        <w:charset w:val="00"/>
        <w:family w:val="Swiss"/>
        <w:notTrueType/>
        <w:pitch w:val="variable"/>
        <w:sig w:usb-0="20000287" w:usb-1="00000001" w:usb-2="00000000" w:usb-3="00000000" w:csb-0="0000019F" w:csb-1="00000000"/>
      </w:font>
      <w:font w:name="European Pi Std 4">
        <w:altName w:val="Calibri"/>
        <w:panose-1 w:val="00000000000000000000"/>
        <w:charset w:val="00"/>
        <w:family w:val="auto"/>
        <w:notTrueType/>
        <w:pitch w:val="default"/>
        <w:sig w:usb-0="00000003" w:usb-1="00000000" w:usb-2="00000000" w:usb-3="00000000" w:csb-0="00000001" w:csb-1="00000000"/>
      </w:font>
    </w:fonts>
  </xsl:param>
  <xsl:param name="lists" as="element(w:lists)">
    <w:lists>
      <w:listDef w:listDefId="0">
        <w:lsid w:val="08F9696C"/>
        <w:plt w:val="HybridMultilevel"/>
        <w:tmpl w:val="23A864EC"/>
        <w:lvl w:ilvl="0" w:tplc="0409000F">
          <w:start w:val="1"/>
          <w:lvlText w:val="%1."/>
          <w:lvlJc w:val="left"/>
          <w:pPr>
            <w:ind w:left="480" w:hanging="480"/>
          </w:pPr>
        </w:lvl>
        <w:lvl w:ilvl="1" w:tplc="04090017">
          <w:start w:val="1"/>
          <w:nfc w:val="20"/>
          <w:lvlText w:val="(%2)"/>
          <w:lvlJc w:val="left"/>
          <w:pPr>
            <w:ind w:left="960" w:hanging="480"/>
          </w:pPr>
        </w:lvl>
        <w:lvl w:ilvl="2" w:tplc="04090011">
          <w:start w:val="1"/>
          <w:nfc w:val="18"/>
          <w:lvlText w:val="%3"/>
          <w:lvlJc w:val="left"/>
          <w:pPr>
            <w:ind w:left="1440" w:hanging="480"/>
          </w:pPr>
        </w:lvl>
        <w:lvl w:ilvl="3" w:tplc="0409000F">
          <w:start w:val="1"/>
          <w:lvlText w:val="%4."/>
          <w:lvlJc w:val="left"/>
          <w:pPr>
            <w:ind w:left="1920" w:hanging="480"/>
          </w:pPr>
        </w:lvl>
        <w:lvl w:ilvl="4" w:tplc="04090017">
          <w:start w:val="1"/>
          <w:nfc w:val="20"/>
          <w:lvlText w:val="(%5)"/>
          <w:lvlJc w:val="left"/>
          <w:pPr>
            <w:ind w:left="2400" w:hanging="480"/>
          </w:pPr>
        </w:lvl>
        <w:lvl w:ilvl="5" w:tplc="04090011">
          <w:start w:val="1"/>
          <w:nfc w:val="18"/>
          <w:lvlText w:val="%6"/>
          <w:lvlJc w:val="left"/>
          <w:pPr>
            <w:ind w:left="2880" w:hanging="480"/>
          </w:pPr>
        </w:lvl>
        <w:lvl w:ilvl="6" w:tplc="0409000F">
          <w:start w:val="1"/>
          <w:lvlText w:val="%7."/>
          <w:lvlJc w:val="left"/>
          <w:pPr>
            <w:ind w:left="3360" w:hanging="480"/>
          </w:pPr>
        </w:lvl>
        <w:lvl w:ilvl="7" w:tplc="04090017">
          <w:start w:val="1"/>
          <w:nfc w:val="20"/>
          <w:lvlText w:val="(%8)"/>
          <w:lvlJc w:val="left"/>
          <w:pPr>
            <w:ind w:left="3840" w:hanging="480"/>
          </w:pPr>
        </w:lvl>
        <w:lvl w:ilvl="8" w:tplc="04090011">
          <w:start w:val="1"/>
          <w:nfc w:val="18"/>
          <w:lvlText w:val="%9"/>
          <w:lvlJc w:val="left"/>
          <w:pPr>
            <w:ind w:left="4320" w:hanging="480"/>
          </w:pPr>
        </w:lvl>
      </w:listDef>
      <w:listDef w:listDefId="1">
        <w:lsid w:val="72254945"/>
        <w:plt w:val="HybridMultilevel"/>
        <w:tmpl w:val="5A606E8E"/>
        <w:lvl w:ilvl="0" w:tplc="04090001">
          <w:start w:val="1"/>
          <w:nfc w:val="23"/>
          <w:lvlText w:val=""/>
          <w:lvlJc w:val="left"/>
          <w:pPr>
            <w:ind w:left="720" w:hanging="360"/>
          </w:pPr>
          <w:rPr>
            <w:rFonts w:ascii="Symbol" w:h-ansi="Symbol" w:hint="default"/>
          </w:rPr>
        </w:lvl>
        <w:lvl w:ilvl="1" w:tplc="04090003" w:tentative="on">
          <w:start w:val="1"/>
          <w:nfc w:val="23"/>
          <w:lvlText w:val="o"/>
          <w:lvlJc w:val="left"/>
          <w:pPr>
            <w:ind w:left="1440" w:hanging="360"/>
          </w:pPr>
          <w:rPr>
            <w:rFonts w:ascii="Courier New" w:h-ansi="Courier New" w:cs="Courier New" w:hint="default"/>
          </w:rPr>
        </w:lvl>
        <w:lvl w:ilvl="2" w:tplc="04090005" w:tentative="on">
          <w:start w:val="1"/>
          <w:nfc w:val="23"/>
          <w:lvlText w:val=""/>
          <w:lvlJc w:val="left"/>
          <w:pPr>
            <w:ind w:left="2160" w:hanging="360"/>
          </w:pPr>
          <w:rPr>
            <w:rFonts w:ascii="Wingdings" w:h-ansi="Wingdings" w:hint="default"/>
          </w:rPr>
        </w:lvl>
        <w:lvl w:ilvl="3" w:tplc="04090001" w:tentative="on">
          <w:start w:val="1"/>
          <w:nfc w:val="23"/>
          <w:lvlText w:val=""/>
          <w:lvlJc w:val="left"/>
          <w:pPr>
            <w:ind w:left="2880" w:hanging="360"/>
          </w:pPr>
          <w:rPr>
            <w:rFonts w:ascii="Symbol" w:h-ansi="Symbol" w:hint="default"/>
          </w:rPr>
        </w:lvl>
        <w:lvl w:ilvl="4" w:tplc="04090003" w:tentative="on">
          <w:start w:val="1"/>
          <w:nfc w:val="23"/>
          <w:lvlText w:val="o"/>
          <w:lvlJc w:val="left"/>
          <w:pPr>
            <w:ind w:left="3600" w:hanging="360"/>
          </w:pPr>
          <w:rPr>
            <w:rFonts w:ascii="Courier New" w:h-ansi="Courier New" w:cs="Courier New" w:hint="default"/>
          </w:rPr>
        </w:lvl>
        <w:lvl w:ilvl="5" w:tplc="04090005" w:tentative="on">
          <w:start w:val="1"/>
          <w:nfc w:val="23"/>
          <w:lvlText w:val=""/>
          <w:lvlJc w:val="left"/>
          <w:pPr>
            <w:ind w:left="4320" w:hanging="360"/>
          </w:pPr>
          <w:rPr>
            <w:rFonts w:ascii="Wingdings" w:h-ansi="Wingdings" w:hint="default"/>
          </w:rPr>
        </w:lvl>
        <w:lvl w:ilvl="6" w:tplc="04090001" w:tentative="on">
          <w:start w:val="1"/>
          <w:nfc w:val="23"/>
          <w:lvlText w:val=""/>
          <w:lvlJc w:val="left"/>
          <w:pPr>
            <w:ind w:left="5040" w:hanging="360"/>
          </w:pPr>
          <w:rPr>
            <w:rFonts w:ascii="Symbol" w:h-ansi="Symbol" w:hint="default"/>
          </w:rPr>
        </w:lvl>
        <w:lvl w:ilvl="7" w:tplc="04090003" w:tentative="on">
          <w:start w:val="1"/>
          <w:nfc w:val="23"/>
          <w:lvlText w:val="o"/>
          <w:lvlJc w:val="left"/>
          <w:pPr>
            <w:ind w:left="5760" w:hanging="360"/>
          </w:pPr>
          <w:rPr>
            <w:rFonts w:ascii="Courier New" w:h-ansi="Courier New" w:cs="Courier New" w:hint="default"/>
          </w:rPr>
        </w:lvl>
        <w:lvl w:ilvl="8" w:tplc="04090005" w:tentative="on">
          <w:start w:val="1"/>
          <w:nfc w:val="23"/>
          <w:lvlText w:val=""/>
          <w:lvlJc w:val="left"/>
          <w:pPr>
            <w:ind w:left="6480" w:hanging="360"/>
          </w:pPr>
          <w:rPr>
            <w:rFonts w:ascii="Wingdings" w:h-ansi="Wingdings" w:hint="default"/>
          </w:rPr>
        </w:lvl>
      </w:listDef>
      <w:listDef w:listDefId="2">
        <w:lsid w:val="78D71815"/>
        <w:plt w:val="Multilevel"/>
        <w:tmpl w:val="BF84A288"/>
        <w:lvl w:ilvl="0">
          <w:start w:val="1"/>
          <w:nfc w:val="1"/>
          <w:pStyle w:val="Heading1"/>
          <w:lvlText w:val="Article %1."/>
          <w:lvlJc w:val="left"/>
          <w:pPr>
            <w:ind w:left="0" w:first-line="0"/>
          </w:pPr>
        </w:lvl>
        <w:lvl w:ilvl="1">
          <w:start w:val="1"/>
          <w:nfc w:val="22"/>
          <w:isLgl/>
          <w:lvlText w:val="Section %1.%2"/>
          <w:lvlJc w:val="left"/>
          <w:pPr>
            <w:ind w:left="0" w:first-line="0"/>
          </w:pPr>
        </w:lvl>
        <w:lvl w:ilvl="2">
          <w:start w:val="1"/>
          <w:nfc w:val="4"/>
          <w:lvlText w:val="(%3)"/>
          <w:lvlJc w:val="left"/>
          <w:pPr>
            <w:ind w:left="720" w:hanging="432"/>
          </w:pPr>
        </w:lvl>
        <w:lvl w:ilvl="3">
          <w:start w:val="1"/>
          <w:nfc w:val="2"/>
          <w:lvlText w:val="(%4)"/>
          <w:lvlJc w:val="right"/>
          <w:pPr>
            <w:ind w:left="864" w:hanging="144"/>
          </w:pPr>
        </w:lvl>
        <w:lvl w:ilvl="4">
          <w:start w:val="1"/>
          <w:lvlText w:val="%5)"/>
          <w:lvlJc w:val="left"/>
          <w:pPr>
            <w:ind w:left="1008" w:hanging="432"/>
          </w:pPr>
        </w:lvl>
        <w:lvl w:ilvl="5">
          <w:start w:val="1"/>
          <w:nfc w:val="4"/>
          <w:lvlText w:val="%6)"/>
          <w:lvlJc w:val="left"/>
          <w:pPr>
            <w:ind w:left="1152" w:hanging="432"/>
          </w:pPr>
        </w:lvl>
        <w:lvl w:ilvl="6">
          <w:start w:val="1"/>
          <w:nfc w:val="2"/>
          <w:lvlText w:val="%7)"/>
          <w:lvlJc w:val="right"/>
          <w:pPr>
            <w:ind w:left="1296" w:hanging="288"/>
          </w:pPr>
        </w:lvl>
        <w:lvl w:ilvl="7">
          <w:start w:val="1"/>
          <w:nfc w:val="4"/>
          <w:lvlText w:val="%8."/>
          <w:lvlJc w:val="left"/>
          <w:pPr>
            <w:ind w:left="1440" w:hanging="432"/>
          </w:pPr>
        </w:lvl>
        <w:lvl w:ilvl="8">
          <w:start w:val="1"/>
          <w:nfc w:val="2"/>
          <w:lvlText w:val="%9."/>
          <w:lvlJc w:val="right"/>
          <w:pPr>
            <w:ind w:left="1584" w:hanging="144"/>
          </w:pPr>
        </w:lvl>
      </w:listDef>
      <w:list w:ilfo="1">
        <w:ilst w:val="2"/>
      </w:list>
      <w:list w:ilfo="2">
        <w:ilst w:val="2"/>
      </w:list>
      <w:list w:ilfo="3">
        <w:ilst w:val="2"/>
      </w:list>
      <w:list w:ilfo="4">
        <w:ilst w:val="2"/>
      </w:list>
      <w:list w:ilfo="5">
        <w:ilst w:val="2"/>
      </w:list>
      <w:list w:ilfo="6">
        <w:ilst w:val="2"/>
      </w:list>
      <w:list w:ilfo="7">
        <w:ilst w:val="2"/>
      </w:list>
      <w:list w:ilfo="8">
        <w:ilst w:val="2"/>
      </w:list>
      <w:list w:ilfo="9">
        <w:ilst w:val="2"/>
      </w:list>
      <w:list w:ilfo="10">
        <w:ilst w:val="0"/>
        <w:lvlOverride w:ilvl="0">
          <w:startOverride w:val="1"/>
        </w:lvlOverride>
        <w:lvlOverride w:ilvl="1">
          <w:startOverride w:val="1"/>
        </w:lvlOverride>
        <w:lvlOverride w:ilvl="2">
          <w:startOverride w:val="1"/>
        </w:lvlOverride>
        <w:lvlOverride w:ilvl="3">
          <w:startOverride w:val="1"/>
        </w:lvlOverride>
        <w:lvlOverride w:ilvl="4">
          <w:startOverride w:val="1"/>
        </w:lvlOverride>
        <w:lvlOverride w:ilvl="5">
          <w:startOverride w:val="1"/>
        </w:lvlOverride>
        <w:lvlOverride w:ilvl="6">
          <w:startOverride w:val="1"/>
        </w:lvlOverride>
        <w:lvlOverride w:ilvl="7">
          <w:startOverride w:val="1"/>
        </w:lvlOverride>
        <w:lvlOverride w:ilvl="8">
          <w:startOverride w:val="1"/>
        </w:lvlOverride>
      </w:list>
      <w:list w:ilfo="11">
        <w:ilst w:val="0"/>
      </w:list>
      <w:list w:ilfo="12">
        <w:ilst w:val="1"/>
      </w:list>
    </w:lists>
  </xsl:param>
  <xsl:param name="styles" as="element(w:styles)">
    <w:styles>
      <w:versionOfBuiltInStylenames w:val="8"/>
      <w:latentStyles w:defLockedState="off" w:latentStyleCount="382">
        <w:lsdException w:name="Normal"/>
        <w:lsdException w:name="heading 1"/>
        <w:lsdException w:name="heading 2"/>
        <w:lsdException w:name="heading 3"/>
        <w:lsdException w:name="heading 4"/>
        <w:lsdException w:name="heading 5"/>
        <w:lsdException w:name="heading 6"/>
        <w:lsdException w:name="heading 7"/>
        <w:lsdException w:name="heading 8"/>
        <w:lsdException w:name="heading 9"/>
        <w:lsdException w:name="index 1"/>
        <w:lsdException w:name="index 2"/>
        <w:lsdException w:name="index 3"/>
        <w:lsdException w:name="index 4"/>
        <w:lsdException w:name="index 5"/>
        <w:lsdException w:name="index 6"/>
        <w:lsdException w:name="index 7"/>
        <w:lsdException w:name="index 8"/>
        <w:lsdException w:name="index 9"/>
        <w:lsdException w:name="toc 1"/>
        <w:lsdException w:name="toc 2"/>
        <w:lsdException w:name="toc 3"/>
        <w:lsdException w:name="toc 4"/>
        <w:lsdException w:name="toc 5"/>
        <w:lsdException w:name="toc 6"/>
        <w:lsdException w:name="toc 7"/>
        <w:lsdException w:name="toc 8"/>
        <w:lsdException w:name="toc 9"/>
        <w:lsdException w:name="Normal Indent"/>
        <w:lsdException w:name="footnote text"/>
        <w:lsdException w:name="annotation text"/>
        <w:lsdException w:name="header"/>
        <w:lsdException w:name="footer"/>
        <w:lsdException w:name="index heading"/>
        <w:lsdException w:name="caption"/>
        <w:lsdException w:name="table of figures"/>
        <w:lsdException w:name="envelope address"/>
        <w:lsdException w:name="envelope return"/>
        <w:lsdException w:name="footnote reference"/>
        <w:lsdException w:name="annotation reference"/>
        <w:lsdException w:name="line number"/>
        <w:lsdException w:name="page number"/>
        <w:lsdException w:name="endnote reference"/>
        <w:lsdException w:name="endnote text"/>
        <w:lsdException w:name="table of authorities"/>
        <w:lsdException w:name="macro"/>
        <w:lsdException w:name="toa heading"/>
        <w:lsdException w:name="List"/>
        <w:lsdException w:name="List Bullet"/>
        <w:lsdException w:name="List Number"/>
        <w:lsdException w:name="List 2"/>
        <w:lsdException w:name="List 3"/>
        <w:lsdException w:name="List 4"/>
        <w:lsdException w:name="List 5"/>
        <w:lsdException w:name="List Bullet 2"/>
        <w:lsdException w:name="List Bullet 3"/>
        <w:lsdException w:name="List Bullet 4"/>
        <w:lsdException w:name="List Bullet 5"/>
        <w:lsdException w:name="List Number 2"/>
        <w:lsdException w:name="List Number 3"/>
        <w:lsdException w:name="List Number 4"/>
        <w:lsdException w:name="List Number 5"/>
        <w:lsdException w:name="Title"/>
        <w:lsdException w:name="Closing"/>
        <w:lsdException w:name="Signature"/>
        <w:lsdException w:name="Default Paragraph Font"/>
        <w:lsdException w:name="Body Text"/>
        <w:lsdException w:name="Body Text Indent"/>
        <w:lsdException w:name="List Continue"/>
        <w:lsdException w:name="List Continue 2"/>
        <w:lsdException w:name="List Continue 3"/>
        <w:lsdException w:name="List Continue 4"/>
        <w:lsdException w:name="List Continue 5"/>
        <w:lsdException w:name="Message Header"/>
        <w:lsdException w:name="Subtitle"/>
        <w:lsdException w:name="Salutation"/>
        <w:lsdException w:name="Date"/>
        <w:lsdException w:name="Body Text First Indent"/>
        <w:lsdException w:name="Body Text First Indent 2"/>
        <w:lsdException w:name="Note Heading"/>
        <w:lsdException w:name="Body Text 2"/>
        <w:lsdException w:name="Body Text 3"/>
        <w:lsdException w:name="Body Text Indent 2"/>
        <w:lsdException w:name="Body Text Indent 3"/>
        <w:lsdException w:name="Block Text"/>
        <w:lsdException w:name="Hyperlink"/>
        <w:lsdException w:name="FollowedHyperlink"/>
        <w:lsdException w:name="Strong"/>
        <w:lsdException w:name="Emphasis"/>
        <w:lsdException w:name="Document Map"/>
        <w:lsdException w:name="Plain Text"/>
        <w:lsdException w:name="E-mail Signature"/>
        <w:lsdException w:name="HTML Top of Form"/>
        <w:lsdException w:name="HTML Bottom of Form"/>
        <w:lsdException w:name="Normal (Web)"/>
        <w:lsdException w:name="HTML Acronym"/>
        <w:lsdException w:name="HTML Address"/>
        <w:lsdException w:name="HTML Cite"/>
        <w:lsdException w:name="HTML Code"/>
        <w:lsdException w:name="HTML Definition"/>
        <w:lsdException w:name="HTML Keyboard"/>
        <w:lsdException w:name="HTML Preformatted"/>
        <w:lsdException w:name="HTML Sample"/>
        <w:lsdException w:name="HTML Typewriter"/>
        <w:lsdException w:name="HTML Variable"/>
        <w:lsdException w:name="Normal Table"/>
        <w:lsdException w:name="annotation subject"/>
        <w:lsdException w:name="No List"/>
        <w:lsdException w:name="Outline List 1"/>
        <w:lsdException w:name="Outline List 2"/>
        <w:lsdException w:name="Outline List 3"/>
        <w:lsdException w:name="Table Simple 1"/>
        <w:lsdException w:name="Table Simple 2"/>
        <w:lsdException w:name="Table Simple 3"/>
        <w:lsdException w:name="Table Classic 1"/>
        <w:lsdException w:name="Table Classic 2"/>
        <w:lsdException w:name="Table Classic 3"/>
        <w:lsdException w:name="Table Classic 4"/>
        <w:lsdException w:name="Table Colorful 1"/>
        <w:lsdException w:name="Table Colorful 2"/>
        <w:lsdException w:name="Table Colorful 3"/>
        <w:lsdException w:name="Table Columns 1"/>
        <w:lsdException w:name="Table Columns 2"/>
        <w:lsdException w:name="Table Columns 3"/>
        <w:lsdException w:name="Table Columns 4"/>
        <w:lsdException w:name="Table Columns 5"/>
        <w:lsdException w:name="Table Grid 1"/>
        <w:lsdException w:name="Table Grid 2"/>
        <w:lsdException w:name="Table Grid 3"/>
        <w:lsdException w:name="Table Grid 4"/>
        <w:lsdException w:name="Table Grid 5"/>
        <w:lsdException w:name="Table Grid 6"/>
        <w:lsdException w:name="Table Grid 7"/>
        <w:lsdException w:name="Table Grid 8"/>
        <w:lsdException w:name="Table List 1"/>
        <w:lsdException w:name="Table List 2"/>
        <w:lsdException w:name="Table List 3"/>
        <w:lsdException w:name="Table List 4"/>
        <w:lsdException w:name="Table List 5"/>
        <w:lsdException w:name="Table List 6"/>
        <w:lsdException w:name="Table List 7"/>
        <w:lsdException w:name="Table List 8"/>
        <w:lsdException w:name="Table 3D effects 1"/>
        <w:lsdException w:name="Table 3D effects 2"/>
        <w:lsdException w:name="Table 3D effects 3"/>
        <w:lsdException w:name="Table Contemporary"/>
        <w:lsdException w:name="Table Elegant"/>
        <w:lsdException w:name="Table Professional"/>
        <w:lsdException w:name="Table Subtle 1"/>
        <w:lsdException w:name="Table Subtle 2"/>
        <w:lsdException w:name="Table Web 1"/>
        <w:lsdException w:name="Table Web 2"/>
        <w:lsdException w:name="Table Web 3"/>
        <w:lsdException w:name="Balloon Text"/>
        <w:lsdException w:name="Table Grid"/>
        <w:lsdException w:name="Table Theme"/>
        <w:lsdException w:name="Note Level 1"/>
        <w:lsdException w:name="Note Level 2"/>
        <w:lsdException w:name="Note Level 3"/>
        <w:lsdException w:name="Note Level 4"/>
        <w:lsdException w:name="Note Level 5"/>
        <w:lsdException w:name="Note Level 6"/>
        <w:lsdException w:name="Note Level 7"/>
        <w:lsdException w:name="Note Level 8"/>
        <w:lsdException w:name="Note Level 9"/>
        <w:lsdException w:name="Placeholder Text"/>
        <w:lsdException w:name="No Spacing"/>
        <w:lsdException w:name="Light Shading"/>
        <w:lsdException w:name="Light List"/>
        <w:lsdException w:name="Light Grid"/>
        <w:lsdException w:name="Medium Shading 1"/>
        <w:lsdException w:name="Medium Shading 2"/>
        <w:lsdException w:name="Medium List 1"/>
        <w:lsdException w:name="Medium List 2"/>
        <w:lsdException w:name="Medium Grid 1"/>
        <w:lsdException w:name="Medium Grid 2"/>
        <w:lsdException w:name="Medium Grid 3"/>
        <w:lsdException w:name="Dark List"/>
        <w:lsdException w:name="Colorful Shading"/>
        <w:lsdException w:name="Colorful List"/>
        <w:lsdException w:name="Colorful Grid"/>
        <w:lsdException w:name="Light Shading Accent 1"/>
        <w:lsdException w:name="Light List Accent 1"/>
        <w:lsdException w:name="Light Grid Accent 1"/>
        <w:lsdException w:name="Medium Shading 1 Accent 1"/>
        <w:lsdException w:name="Medium Shading 2 Accent 1"/>
        <w:lsdException w:name="Medium List 1 Accent 1"/>
        <w:lsdException w:name="Revision"/>
        <w:lsdException w:name="List Paragraph"/>
        <w:lsdException w:name="Quote"/>
        <w:lsdException w:name="Intense Quote"/>
        <w:lsdException w:name="Medium List 2 Accent 1"/>
        <w:lsdException w:name="Medium Grid 1 Accent 1"/>
        <w:lsdException w:name="Medium Grid 2 Accent 1"/>
        <w:lsdException w:name="Medium Grid 3 Accent 1"/>
        <w:lsdException w:name="Dark List Accent 1"/>
        <w:lsdException w:name="Colorful Shading Accent 1"/>
        <w:lsdException w:name="Colorful List Accent 1"/>
        <w:lsdException w:name="Colorful Grid Accent 1"/>
        <w:lsdException w:name="Light Shading Accent 2"/>
        <w:lsdException w:name="Light List Accent 2"/>
        <w:lsdException w:name="Light Grid Accent 2"/>
        <w:lsdException w:name="Medium Shading 1 Accent 2"/>
        <w:lsdException w:name="Medium Shading 2 Accent 2"/>
        <w:lsdException w:name="Medium List 1 Accent 2"/>
        <w:lsdException w:name="Medium List 2 Accent 2"/>
        <w:lsdException w:name="Medium Grid 1 Accent 2"/>
        <w:lsdException w:name="Medium Grid 2 Accent 2"/>
        <w:lsdException w:name="Medium Grid 3 Accent 2"/>
        <w:lsdException w:name="Dark List Accent 2"/>
        <w:lsdException w:name="Colorful Shading Accent 2"/>
        <w:lsdException w:name="Colorful List Accent 2"/>
        <w:lsdException w:name="Colorful Grid Accent 2"/>
        <w:lsdException w:name="Light Shading Accent 3"/>
        <w:lsdException w:name="Light List Accent 3"/>
        <w:lsdException w:name="Light Grid Accent 3"/>
        <w:lsdException w:name="Medium Shading 1 Accent 3"/>
        <w:lsdException w:name="Medium Shading 2 Accent 3"/>
        <w:lsdException w:name="Medium List 1 Accent 3"/>
        <w:lsdException w:name="Medium List 2 Accent 3"/>
        <w:lsdException w:name="Medium Grid 1 Accent 3"/>
        <w:lsdException w:name="Medium Grid 2 Accent 3"/>
        <w:lsdException w:name="Medium Grid 3 Accent 3"/>
        <w:lsdException w:name="Dark List Accent 3"/>
        <w:lsdException w:name="Colorful Shading Accent 3"/>
        <w:lsdException w:name="Colorful List Accent 3"/>
        <w:lsdException w:name="Colorful Grid Accent 3"/>
        <w:lsdException w:name="Light Shading Accent 4"/>
        <w:lsdException w:name="Light List Accent 4"/>
        <w:lsdException w:name="Light Grid Accent 4"/>
        <w:lsdException w:name="Medium Shading 1 Accent 4"/>
        <w:lsdException w:name="Medium Shading 2 Accent 4"/>
        <w:lsdException w:name="Medium List 1 Accent 4"/>
        <w:lsdException w:name="Medium List 2 Accent 4"/>
        <w:lsdException w:name="Medium Grid 1 Accent 4"/>
        <w:lsdException w:name="Medium Grid 2 Accent 4"/>
        <w:lsdException w:name="Medium Grid 3 Accent 4"/>
        <w:lsdException w:name="Dark List Accent 4"/>
        <w:lsdException w:name="Colorful Shading Accent 4"/>
        <w:lsdException w:name="Colorful List Accent 4"/>
        <w:lsdException w:name="Colorful Grid Accent 4"/>
        <w:lsdException w:name="Light Shading Accent 5"/>
        <w:lsdException w:name="Light List Accent 5"/>
        <w:lsdException w:name="Light Grid Accent 5"/>
        <w:lsdException w:name="Medium Shading 1 Accent 5"/>
        <w:lsdException w:name="Medium Shading 2 Accent 5"/>
        <w:lsdException w:name="Medium List 1 Accent 5"/>
        <w:lsdException w:name="Medium List 2 Accent 5"/>
        <w:lsdException w:name="Medium Grid 1 Accent 5"/>
        <w:lsdException w:name="Medium Grid 2 Accent 5"/>
        <w:lsdException w:name="Medium Grid 3 Accent 5"/>
        <w:lsdException w:name="Dark List Accent 5"/>
        <w:lsdException w:name="Colorful Shading Accent 5"/>
        <w:lsdException w:name="Colorful List Accent 5"/>
        <w:lsdException w:name="Colorful Grid Accent 5"/>
        <w:lsdException w:name="Light Shading Accent 6"/>
        <w:lsdException w:name="Light List Accent 6"/>
        <w:lsdException w:name="Light Grid Accent 6"/>
        <w:lsdException w:name="Medium Shading 1 Accent 6"/>
        <w:lsdException w:name="Medium Shading 2 Accent 6"/>
        <w:lsdException w:name="Medium List 1 Accent 6"/>
        <w:lsdException w:name="Medium List 2 Accent 6"/>
        <w:lsdException w:name="Medium Grid 1 Accent 6"/>
        <w:lsdException w:name="Medium Grid 2 Accent 6"/>
        <w:lsdException w:name="Medium Grid 3 Accent 6"/>
        <w:lsdException w:name="Dark List Accent 6"/>
        <w:lsdException w:name="Colorful Shading Accent 6"/>
        <w:lsdException w:name="Colorful List Accent 6"/>
        <w:lsdException w:name="Colorful Grid Accent 6"/>
        <w:lsdException w:name="Subtle Emphasis"/>
        <w:lsdException w:name="Intense Emphasis"/>
        <w:lsdException w:name="Subtle Reference"/>
        <w:lsdException w:name="Intense Reference"/>
        <w:lsdException w:name="Book Title"/>
        <w:lsdException w:name="Bibliography"/>
        <w:lsdException w:name="TOC Heading"/>
        <w:lsdException w:name="Plain Table 1"/>
        <w:lsdException w:name="Plain Table 2"/>
        <w:lsdException w:name="Plain Table 3"/>
        <w:lsdException w:name="Plain Table 4"/>
        <w:lsdException w:name="Plain Table 5"/>
        <w:lsdException w:name="Grid Table Light"/>
        <w:lsdException w:name="Grid Table 1 Light"/>
        <w:lsdException w:name="Grid Table 2"/>
        <w:lsdException w:name="Grid Table 3"/>
        <w:lsdException w:name="Grid Table 4"/>
        <w:lsdException w:name="Grid Table 5 Dark"/>
        <w:lsdException w:name="Grid Table 6 Colorful"/>
        <w:lsdException w:name="Grid Table 7 Colorful"/>
        <w:lsdException w:name="Grid Table 1 Light Accent 1"/>
        <w:lsdException w:name="Grid Table 2 Accent 1"/>
        <w:lsdException w:name="Grid Table 3 Accent 1"/>
        <w:lsdException w:name="Grid Table 4 Accent 1"/>
        <w:lsdException w:name="Grid Table 5 Dark Accent 1"/>
        <w:lsdException w:name="Grid Table 6 Colorful Accent 1"/>
        <w:lsdException w:name="Grid Table 7 Colorful Accent 1"/>
        <w:lsdException w:name="Grid Table 1 Light Accent 2"/>
        <w:lsdException w:name="Grid Table 2 Accent 2"/>
        <w:lsdException w:name="Grid Table 3 Accent 2"/>
        <w:lsdException w:name="Grid Table 4 Accent 2"/>
        <w:lsdException w:name="Grid Table 5 Dark Accent 2"/>
        <w:lsdException w:name="Grid Table 6 Colorful Accent 2"/>
        <w:lsdException w:name="Grid Table 7 Colorful Accent 2"/>
        <w:lsdException w:name="Grid Table 1 Light Accent 3"/>
        <w:lsdException w:name="Grid Table 2 Accent 3"/>
        <w:lsdException w:name="Grid Table 3 Accent 3"/>
        <w:lsdException w:name="Grid Table 4 Accent 3"/>
        <w:lsdException w:name="Grid Table 5 Dark Accent 3"/>
        <w:lsdException w:name="Grid Table 6 Colorful Accent 3"/>
        <w:lsdException w:name="Grid Table 7 Colorful Accent 3"/>
        <w:lsdException w:name="Grid Table 1 Light Accent 4"/>
        <w:lsdException w:name="Grid Table 2 Accent 4"/>
        <w:lsdException w:name="Grid Table 3 Accent 4"/>
        <w:lsdException w:name="Grid Table 4 Accent 4"/>
        <w:lsdException w:name="Grid Table 5 Dark Accent 4"/>
        <w:lsdException w:name="Grid Table 6 Colorful Accent 4"/>
        <w:lsdException w:name="Grid Table 7 Colorful Accent 4"/>
        <w:lsdException w:name="Grid Table 1 Light Accent 5"/>
        <w:lsdException w:name="Grid Table 2 Accent 5"/>
        <w:lsdException w:name="Grid Table 3 Accent 5"/>
        <w:lsdException w:name="Grid Table 4 Accent 5"/>
        <w:lsdException w:name="Grid Table 5 Dark Accent 5"/>
        <w:lsdException w:name="Grid Table 6 Colorful Accent 5"/>
        <w:lsdException w:name="Grid Table 7 Colorful Accent 5"/>
        <w:lsdException w:name="Grid Table 1 Light Accent 6"/>
        <w:lsdException w:name="Grid Table 2 Accent 6"/>
        <w:lsdException w:name="Grid Table 3 Accent 6"/>
        <w:lsdException w:name="Grid Table 4 Accent 6"/>
        <w:lsdException w:name="Grid Table 5 Dark Accent 6"/>
        <w:lsdException w:name="Grid Table 6 Colorful Accent 6"/>
        <w:lsdException w:name="Grid Table 7 Colorful Accent 6"/>
        <w:lsdException w:name="List Table 1 Light"/>
        <w:lsdException w:name="List Table 2"/>
        <w:lsdException w:name="List Table 3"/>
        <w:lsdException w:name="List Table 4"/>
        <w:lsdException w:name="List Table 5 Dark"/>
        <w:lsdException w:name="List Table 6 Colorful"/>
        <w:lsdException w:name="List Table 7 Colorful"/>
        <w:lsdException w:name="List Table 1 Light Accent 1"/>
        <w:lsdException w:name="List Table 2 Accent 1"/>
        <w:lsdException w:name="List Table 3 Accent 1"/>
        <w:lsdException w:name="List Table 4 Accent 1"/>
        <w:lsdException w:name="List Table 5 Dark Accent 1"/>
        <w:lsdException w:name="List Table 6 Colorful Accent 1"/>
        <w:lsdException w:name="List Table 7 Colorful Accent 1"/>
        <w:lsdException w:name="List Table 1 Light Accent 2"/>
        <w:lsdException w:name="List Table 2 Accent 2"/>
        <w:lsdException w:name="List Table 3 Accent 2"/>
        <w:lsdException w:name="List Table 4 Accent 2"/>
        <w:lsdException w:name="List Table 5 Dark Accent 2"/>
        <w:lsdException w:name="List Table 6 Colorful Accent 2"/>
        <w:lsdException w:name="List Table 7 Colorful Accent 2"/>
        <w:lsdException w:name="List Table 1 Light Accent 3"/>
        <w:lsdException w:name="List Table 2 Accent 3"/>
        <w:lsdException w:name="List Table 3 Accent 3"/>
        <w:lsdException w:name="List Table 4 Accent 3"/>
        <w:lsdException w:name="List Table 5 Dark Accent 3"/>
        <w:lsdException w:name="List Table 6 Colorful Accent 3"/>
        <w:lsdException w:name="List Table 7 Colorful Accent 3"/>
        <w:lsdException w:name="List Table 1 Light Accent 4"/>
        <w:lsdException w:name="List Table 2 Accent 4"/>
        <w:lsdException w:name="List Table 3 Accent 4"/>
        <w:lsdException w:name="List Table 4 Accent 4"/>
        <w:lsdException w:name="List Table 5 Dark Accent 4"/>
        <w:lsdException w:name="List Table 6 Colorful Accent 4"/>
        <w:lsdException w:name="List Table 7 Colorful Accent 4"/>
        <w:lsdException w:name="List Table 1 Light Accent 5"/>
        <w:lsdException w:name="List Table 2 Accent 5"/>
        <w:lsdException w:name="List Table 3 Accent 5"/>
        <w:lsdException w:name="List Table 4 Accent 5"/>
        <w:lsdException w:name="List Table 5 Dark Accent 5"/>
        <w:lsdException w:name="List Table 6 Colorful Accent 5"/>
        <w:lsdException w:name="List Table 7 Colorful Accent 5"/>
        <w:lsdException w:name="List Table 1 Light Accent 6"/>
        <w:lsdException w:name="List Table 2 Accent 6"/>
        <w:lsdException w:name="List Table 3 Accent 6"/>
        <w:lsdException w:name="List Table 4 Accent 6"/>
        <w:lsdException w:name="List Table 5 Dark Accent 6"/>
        <w:lsdException w:name="List Table 6 Colorful Accent 6"/>
        <w:lsdException w:name="List Table 7 Colorful Accent 6"/>
        <w:lsdException w:name="Mention"/>
        <w:lsdException w:name="Smart Hyperlink"/>
      </w:latentStyles>
      <w:style w:type="paragraph" w:default="on" w:styleId="Normal">
        <w:name w:val="Normal"/>
        <w:rsid w:val="00B65110"/>
        <w:pPr>
          <w:supressLineNumbers/>
        </w:pPr>
        <w:rPr>
          <wx:font wx:val="Calibri"/>
          <w:sz w:val="22"/>
          <w:sz-cs w:val="22"/>
          <w:lang w:val="EN-US" w:fareast="EN-US" w:bidi="AR-SA"/>
        </w:rPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="Heading1">
        <w:name w:val="heading 1"/>
        <wx:uiName wx:val="Heading 1"/>
        <w:basedOn w:val="Normal"/>
        <w:next w:val="Normal"/>
        <w:link w:val="Heading1Char"/>
        <w:rsid w:val="00655D3F"/>
        <w:pPr>
          <w:keepNext/>
          <w:keepLines/>
          <w:listPr>
            <w:ilfo w:val="1"/>
          </w:listPr>
          <w:spacing w:before="240"/>
        </w:pPr>
        <w:rPr>
          <w:rFonts w:ascii="Cambria" w:fareast="Times New Roman" w:h-ansi="Cambria"/>
          <wx:font wx:val="Cambria"/>
          <w:color w:val="365F91"/>
          <w:sz w:val="32"/>
          <w:sz-cs w:val="32"/>
        </w:rPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="Heading2">
        <w:name w:val="heading 2"/>
        <wx:uiName wx:val="Heading 2"/>
        <w:basedOn w:val="Normal"/>
        <w:next w:val="Normal"/>
        <w:link w:val="Heading2Char"/>
        <w:rsid w:val="00655D3F"/>
        <w:pPr>
          <w:keepNext/>
          <w:keepLines/>
          <w:spacing w:before="200"/>
        </w:pPr>
        <w:rPr>
          <w:rFonts w:ascii="Cambria" w:fareast="Times New Roman" w:h-ansi="Cambria"/>
          <wx:font wx:val="Cambria"/>
          <w:b/>
          <w:b-cs/>
          <w:color w:val="4F81BD"/>
          <w:sz w:val="26"/>
          <w:sz-cs w:val="26"/>
        </w:rPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="Heading3">
        <w:name w:val="heading 3"/>
        <wx:uiName wx:val="Heading 3"/>
        <w:basedOn w:val="Normal"/>
        <w:next w:val="Normal"/>
        <w:link w:val="Heading3Char"/>
        <w:rsid w:val="00655D3F"/>
        <w:pPr>
          <w:keepNext/>
          <w:keepLines/>
          <w:spacing w:before="200"/>
        </w:pPr>
        <w:rPr>
          <w:rFonts w:ascii="Cambria" w:fareast="Times New Roman" w:h-ansi="Cambria"/>
          <wx:font wx:val="Cambria"/>
          <w:b/>
          <w:b-cs/>
          <w:color w:val="4F81BD"/>
        </w:rPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="Heading4">
        <w:name w:val="heading 4"/>
        <wx:uiName wx:val="Heading 4"/>
        <w:basedOn w:val="Normal"/>
        <w:next w:val="Normal"/>
        <w:link w:val="Heading4Char"/>
        <w:rsid w:val="00B86518"/>
        <w:pPr>
          <w:keepNext/>
          <w:keepLines/>
          <w:spacing w:before="40"/>
        </w:pPr>
        <w:rPr>
          <w:rFonts w:ascii="Cambria" w:fareast="Times New Roman" w:h-ansi="Cambria"/>
          <wx:font wx:val="Cambria"/>
          <w:i/>
          <w:i-cs/>
          <w:color w:val="365F91"/>
        </w:rPr>
      </w:style>
      <w:style w:type="character" w:default="on" w:styleId="DefaultParagraphFont">
        <w:name w:val="Default Paragraph Font"/>
      </w:style>
      <w:style w:type="table" w:default="on" w:styleId="TableNormal">
        <w:name w:val="Normal Table"/>
        <wx:uiName wx:val="Table Normal"/>
        <w:rPr>
          <wx:font wx:val="Calibri"/>
          <w:lang w:val="EN-US" w:fareast="EN-US" w:bidi="AR-SA"/>
        </w:rPr>
        <w:tblPr>
          <w:tblInd w:w="0" w:type="dxa"/>
          <w:tblCellMar>
            <w:top w:w="0" w:type="dxa"/>
            <w:left w:w="108" w:type="dxa"/>
            <w:bottom w:w="0" w:type="dxa"/>
            <w:right w:w="108" w:type="dxa"/>
          </w:tblCellMar>
        </w:tblPr>
      </w:style>
      <w:style w:type="list" w:default="on" w:styleId="NoList">
        <w:name w:val="No List"/>
      </w:style>
      <xsl:copy-of select="$xspec-styles"/>
    </w:styles>
  </xsl:param>
  
  <xsl:param name="xspec-styles">
    <xsl:comment select="'Paragraph styles used by xspec'"/>
    <w:style w:type="paragraph" w:styleId="x:params">
      <w:name w:val="x:params"/>
      <w:pPr>
        <w:pBdr>
          <w:top w:val="single" w:sz="18" wx:bdrwidth="45" w:space="1" w:color="4472C4"/>
          <w:left w:val="single" w:sz="18" wx:bdrwidth="45" w:space="4" w:color="4472C4"/>
          <w:bottom w:val="single" w:sz="18" wx:bdrwidth="45" w:space="1" w:color="4472C4"/>
          <w:right w:val="single" w:sz="18" wx:bdrwidth="45" w:space="4" w:color="4472C4"/>
        </w:pBdr>
        <w:shd w:val="clear" w:color="auto" w:fill="E7E6E6"/>
        <w:tabs>
          <w:tab w:val="right" w:pos="9000"/>
        </w:tabs>
        <w:spacing w:before="240" w:after="240"/>
        <w:contextualSpacing/>
      </w:pPr>
      <w:rPr>
        <wx:font wx:val="Calibri"/>
        <w:sz w:val="32"/>
        <w:lang w:val="EN-US" w:fareast="EN-US" w:bidi="AR-SA"/>
      </w:rPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="x:scenario-start">
      <w:name w:val="x:scenario-start"/>
      <w:pPr>
        <w:keepNext w:val="on"/>
        <w:pBdr>
          <w:top w:val="single" w:sz="18" wx:bdrwidth="45" w:space="1" w:color="70AD47"/>
          <w:left w:val="single" w:sz="18" wx:bdrwidth="45" w:space="4" w:color="70AD47"/>
          <w:right w:val="single" w:sz="18" wx:bdrwidth="45" w:space="4" w:color="70AD47"/>
        </w:pBdr>
        <w:shd w:val="clear" w:color="auto" w:fill="E7E6E6"/>
        <w:tabs>
          <w:tab w:val="right" w:pos="9000"/>
        </w:tabs>
        <w:spacing w:before="600"/>
      </w:pPr>
      <w:rPr>
        <wx:font wx:val="Calibri"/>
        <w:sz w:val="32"/>
      </w:rPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="x:scenario-end">
      <w:name w:val="x:scenario-end"/>
      <w:basedOn w:val="x:scenario-start"/>
      <w:pPr>
        <w:keepNext w:val="off"/>
        <w:pBdr>
          <w:top w:val="none" w:sz="0" wx:bdrwidth="0" w:space="0" w:color="auto"/>
          <w:bottom w:val="single" w:sz="18" wx:bdrwidth="45" w:space="1" w:color="70AD47"/>
        </w:pBdr>
        <w:spacing w:before="0" w:after="1800"/>
        <w:contextualSpacing/>
      </w:pPr>
      <w:rPr>
        <wx:font wx:val="Calibri"/>
      </w:rPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="x:scenario-pending-start">
      <w:name w:val="x:scenario-pending-start"/>
      <w:basedOn w:val="x:scenario-start"/>
      <w:pPr>
        <w:pBdr>
          <w:top w:val="single" w:sz="18" wx:bdrwidth="45" w:space="1" w:color="767171"/>
          <w:left w:val="single" w:sz="18" wx:bdrwidth="45" w:space="4" w:color="767171"/>
          <w:right w:val="single" w:sz="18" wx:bdrwidth="45" w:space="4" w:color="767171"/>
        </w:pBdr>
      </w:pPr>
      <w:rPr>
        <wx:font wx:val="Calibri"/>
      </w:rPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="x:pending-end">
      <w:name w:val="x:pending-end"/>
      <w:basedOn w:val="x:pending"/>
      <w:pPr>
        <w:keepNext w:val="off"/>
        <w:pBdr>
          <w:top w:val="none" w:sz="0" wx:bdrwidth="0" w:space="0" w:color="auto"/>
          <w:bottom w:val="single" w:sz="18" wx:bdrwidth="45" w:space="1" w:color="767171"/>
        </w:pBdr>
        <w:spacing w:after="240"/>
      </w:pPr>
      <w:rPr>
        <wx:font wx:val="Calibri"/>
      </w:rPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="x:scenario-pending-end">
      <w:name w:val="x:scenario-pending-end"/>
      <w:basedOn w:val="x:scenario-pending"/>
      <w:pPr>
        <w:keepNext w:val="off"/>
        <w:pBdr>
          <w:top w:val="none" w:sz="0" wx:bdrwidth="0" w:space="0" w:color="auto"/>
          <w:bottom w:val="single" w:sz="18" wx:bdrwidth="45" w:space="1" w:color="767171"/>
        </w:pBdr>
        <w:spacing w:before="0" w:after="1800"/>
        <w:contextualSpacing/>
      </w:pPr>
      <w:rPr>
        <wx:font wx:val="Calibri"/>
      </w:rPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="x:scenario-pending">
      <w:name w:val="x:scenario-pending"/>
      <w:basedOn w:val="x:scenario-start"/>
      <w:rsid w:val="00547639"/>
      <w:pPr>
        <w:pBdr>
          <w:top w:val="single" w:sz="18" wx:bdrwidth="45" w:space="1" w:color="767171"/>
          <w:left w:val="single" w:sz="18" wx:bdrwidth="45" w:space="4" w:color="767171"/>
          <w:right w:val="single" w:sz="18" wx:bdrwidth="45" w:space="4" w:color="767171"/>
        </w:pBdr>
        <w:spacing w:before="0"/>
      </w:pPr>
      <w:rPr>
        <wx:font wx:val="Calibri"/>
      </w:rPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="x:context">
      <w:name w:val="x:context"/>
      <w:pPr>
        <w:keepNext/>
        <w:pBdr>
          <w:bottom w:val="single" w:sz="24" wx:bdrwidth="60" w:space="1" w:color="auto"/>
        </w:pBdr>
        <w:shd w:val="clear" w:color="auto" w:fill="E7E6E6"/>
        <w:tabs>
          <w:tab w:val="right" w:pos="9000"/>
        </w:tabs>
        <w:spacing w:before="240" w:after="120"/>
      </w:pPr>
      <w:rPr>
        <wx:font wx:val="Calibri"/>
        <w:sz w:val="24"/>
      </w:rPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="x:context-end">
      <w:name w:val="x:context-end"/>
      <w:basedOn w:val="x:context"/>
      <w:pPr>
        <w:keepNext w:val="off"/>
        <w:pBdr>
          <w:top w:val="single" w:sz="24" wx:bdrwidth="60" w:space="1" w:color="auto"/>
          <w:bottom w:val="none" w:sz="0" wx:bdrwidth="0" w:space="0" w:color="auto"/>
        </w:pBdr>
        <w:spacing w:before="120" w:after="240"/>
      </w:pPr>
      <w:rPr>
        <wx:font wx:val="Calibri"/>
      </w:rPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="x:call-template">
      <w:name w:val="x:call-template"/>
      <w:basedOn w:val="x:expect"/>
      <w:pPr>
        <w:pBdr>
          <w:bottom w:val="dashed" w:sz="4" wx:bdrwidth="10" w:space="4" w:color="auto"/>
        </w:pBdr>
        <w:contextualSpacing/>
      </w:pPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="x:call-function">
      <w:name w:val="x:call-function"/>
      <w:basedOn w:val="x:call-template"/>
    </w:style>
    <w:style w:type="paragraph" w:styleId="x:expect">
      <w:name w:val="x:expect"/>
      <w:basedOn w:val="x:context"/>
      <w:pPr>
        <w:keepNext w:val="on"/>
        <w:pBdr>
          <w:top w:val="dashed" w:sz="4" wx:bdrwidth="10" w:space="1" w:color="auto"/>
          <w:left w:val="dashed" w:sz="4" wx:bdrwidth="10" w:space="4" w:color="auto"/>
          <w:bottom w:val="none" w:sz="0" wx:bdrwidth="0" w:space="0" w:color="auto"/>
          <w:right w:val="dashed" w:sz="4" wx:bdrwidth="10" w:space="4" w:color="auto"/>
        </w:pBdr>
        <w:spacing w:after="0"/>
      </w:pPr>
      <w:rPr>
        <wx:font wx:val="Calibri"/>
      </w:rPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="x:expect-details">
      <w:name w:val="x:expect-details"/>
      <w:basedOn w:val="x:expect"/>
      <w:pPr>
        <w:pBdr>
          <w:top w:val="none" w:sz="0" wx:bdrwidth="0" w:space="0" w:color="auto"/>
        </w:pBdr>
        <w:spacing w:before="0"/>
      </w:pPr>
      <w:rPr>
        <wx:font wx:val="Calibri"/>
      </w:rPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="x:expect-end">
      <w:name w:val="x:expect-end"/>
      <w:basedOn w:val="x:expect"/>
      <w:pPr>
        <w:keepNext w:val="off"/>
        <w:pBdr>
          <w:top w:val="none" w:sz="0" wx:bdrwidth="0" w:space="0" w:color="auto"/>
          <w:bottom w:val="dashed" w:sz="8" wx:bdrwidth="20" w:space="1" w:color="auto"/>
        </w:pBdr>
        <w:spacing w:before="0" w:after="120"/>
      </w:pPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="x:pending">
      <w:name w:val="x:pending"/>
      <w:basedOn w:val="x:scenario-pending"/>
    </w:style>
    <w:style w:type="paragraph" w:styleId="x:code">
      <w:name w:val="x:code"/>
      <w:pPr>
        <w:pBdr>
          <w:top w:val="single" w:sz="4" wx:bdrwidth="20" w:space="1" w:color="auto" w:shadow="on"/>
          <w:left w:val="single" w:sz="4" wx:bdrwidth="20" w:space="4" w:color="auto" w:shadow="on"/>
          <w:bottom w:val="single" w:sz="4" wx:bdrwidth="20" w:space="1" w:color="auto" w:shadow="on"/>
          <w:right w:val="single" w:sz="4" wx:bdrwidth="20" w:space="4" w:color="auto" w:shadow="on"/>
        </w:pBdr>
        <w:spacing w:before="120" w:after="120"/>
        <w:ind w:left="284" w:right="284"/>
        <w:contextualSpacing/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Courier" w:h-ansi="Courier"/>
        <wx:font wx:val="Courier"/>
        <w:noProof/>
      </w:rPr>
      <w:basedOn w:val="normal"/>
    </w:style>
    <xsl:comment select="'Character styles used by xspec'"/>
    <w:style w:type="character" w:styleId="x:container-label">
      <w:name w:val="x:container-label"/>
      <w:rPr>
        <w:b/>
        <w:smallCaps/>
        <w:shadow/>
        <w:sz w:val="32"/>
        <w:bdr w:val="none" w:sz="0" wx:bdrwidth="0" w:space="0" w:color="auto"/>
        <w:shd w:val="clear" w:color="auto" w:fill="auto"/>
      </w:rPr>
    </w:style>
    <w:style w:type="character" w:styleId="x:inline-code">
      <w:name w:val="x:inline-code"/>
      <w:rPr>
        <w:rFonts w:ascii="Courier" w:h-ansi="Courier"/>
        <w:noProof/>
        <w:bdr w:val="single" w:sz="4" wx:bdrwidth="10" w:space="0" w:color="B4C6E7"/>
        <w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/>
      </w:rPr>
    </w:style>
    <w:style w:type="character" w:styleId="x:code-element-furniture">
      <w:name w:val="x:code-element-furniture"/>
      <w:rPr>
        <w:color w:val="0070C0"/>
        <w:bdr w:val="none" w:sz="0" wx:bdrwidth="0" w:space="0" w:color="auto"/>
        <w:shd w:val="clear" w:color="auto" w:fill="auto"/>
      </w:rPr>
      <w:basedOn w:val="x:inline-code"/>
    </w:style>
    <w:style w:type="character" w:styleId="x:code-element-name">
      <w:name w:val="x:code-element-name"/>
      <w:rPr>
        <w:color w:val="00B0F0"/>
      </w:rPr>
      <w:basedOn w:val="x:code-element-furniture"/>
    </w:style>
    <w:style w:type="character" w:styleId="x:code-attribute-name">
      <w:name w:val="x:code-attribute-name"/>
      <w:rPr>
        <w:color w:val="FF0000"/>
      </w:rPr>
      <w:basedOn w:val="x:code-element-furniture"/>
    </w:style>
    <w:style w:type="character" w:styleId="x:code-attribute-furniture">
      <w:name w:val="x:code-attribute-furniture"/>
      <w:rPr>
        <w:color w:val="0070C0"/>
      </w:rPr>
      <w:basedOn w:val="x:code-element-furniture"/>
    </w:style>
    <w:style w:type="character" w:styleId="x:code-attribute-content">
      <w:name w:val="x:code-attribute-content"/>
      <w:rPr>
        <w:color w:val="ED7D31"/>
      </w:rPr>
      <w:basedOn w:val="x:code-element-furniture"/>
    </w:style>
    <w:style w:type="character" w:styleId="x:code-comment">
      <w:name w:val="x:code-comment"/>
      <w:rPr>
        <w:color w:val="808080"/>
      </w:rPr>
      <w:basedOn w:val="x:code-element-furniture"/>
    </w:style>
    <w:style w:type="character" w:styleId="x:code-pi">
      <w:name w:val="x:code-pi"/>
      <w:rPr>
        <w:color w:val="92D050"/>
      </w:rPr>
      <w:basedOn w:val="x:code-element-furniture"/>
    </w:style>
  </xsl:param>

  <xsl:template match="node() | @* | document[x2w:isWord(.)] | body[x2w:isWord(.)]" priority="-0.6">
    <xsl:apply-templates select="@*, node()"/>
  </xsl:template>

  <!-- Copy word elements; Change namespace to friendlier 2003 format -->
  <xsl:template match="*[x2w:isWord(.)]" priority="-0.55">
    <xsl:element name="w:{local-name()}">
      <xsl:apply-templates select="@*, node()"/>
    </xsl:element>
  </xsl:template>

  <!-- Wrap runs in paragraphs -->
  <xsl:template match="*:r[x2w:isWord(.) and not(ancestor::*:p[x2w:isWord(.)])]">
    <w:p>
      <w:r>
        <xsl:apply-templates select="@*, node()"/>
      </w:r>
    </w:p>
  </xsl:template>

  <!-- Don't forget to output word text -->
  <xsl:template match="*:t[x2w:isWord(.)]/text()">
    <xsl:attribute name="xml:space" select="'preserve'"/>
    <xsl:copy/>
  </xsl:template>

  <xsl:template match="@*[x2w:isWord(.)]">
    <xsl:attribute name="w:{local-name()}" select="."/>
  </xsl:template>

  <!-- Boilerplate for word document, styles etc -->
  <xsl:template match="/x:description">
    <xsl:processing-instruction name="mso-application">progid="Word.Document"</xsl:processing-instruction>
    <w:wordDocument w:macrosPresent="no" w:embeddedObjPresent="no" w:ocxPresent="no">
      <xsl:apply-templates mode="style" select="($fonts, $lists, $styles)"/>
      <w:body>
        <xsl:apply-templates select="node()"/>
      </w:body>
    </w:wordDocument>
  </xsl:template>
  
  <xsl:mode name="style" on-no-match="shallow-copy"/>
  
  <xsl:template match="w:style[@w:styleId = $xspec-styles/w:style/@w:styleId]" mode="style"/>
  
  <xsl:template match="*:outlineLvl[x2w:isWord(.)]" mode="style #default"/>
  
  <xsl:template match="w:styles" mode="style">
    <xsl:copy>
      <xsl:apply-templates select="@*, node()" mode="style"/>
      <xsl:copy-of select="$xspec-styles"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="x:description/x:param">
    <w:p>
      <w:pPr>
        <w:pStyle w:val="x:params"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rStyle w:val="x:inline-code"/>
        </w:rPr>
        <w:t>${@name}</w:t>
      </w:r>
      <w:r>
        <w:t xml:space="preserve"> has value: </w:t>
      </w:r>
      <xsl:apply-templates select="@select"/>
      <w:r>
        <w:tab/>
      </w:r>
      <w:r>
        <w:rPr>
          <w:rStyle w:val="x:inline-code"/>
        </w:rPr>
        <w:t>{@as}</w:t>
      </w:r>
      <w:r>
        <xsl:if test="@as">
          <w:t xml:space="preserve"> type </w:t>
        </xsl:if>
      </w:r>
      <w:r>
        <w:rPr>
          <w:rStyle w:val="x:container-label"/>
        </w:rPr>
        <w:t>Global Parameter</w:t>
      </w:r>
    </w:p>
    <xsl:apply-templates select="@href">
      <xsl:with-param name="pStyle" select="'x:params'"/>
    </xsl:apply-templates>
    <xsl:call-template name="xml"/>
  </xsl:template>

  <xsl:template match="x:scenario[not(@pending)]">
    <xsl:call-template name="x:scenario">
      <xsl:with-param name="style" select="'x:scenario'"/>
    </xsl:call-template>
  </xsl:template>

  <xsl:template match="x:scenario[@pending]">
    <xsl:call-template name="x:scenario">
      <xsl:with-param name="style" select="'x:scenario-pending'"/>
    </xsl:call-template>
  </xsl:template>

  <xsl:template name="x:scenario">
    <xsl:param name="style" required="yes" as="xs:string"/>
    <xsl:param name="ScenarioLevel" tunnel="yes" as="xs:integer" select="0"/>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="{$style}-start"/>
        <w:outlineLvl w:val="{$ScenarioLevel}"/>
      </w:pPr>
      <w:r>
        <w:t>
          <xsl:value-of select="@label"/>
        </w:t>
      </w:r>
      <w:r>
        <w:rPr>
          <w:rStyle w:val="x:container-label"/>
        </w:rPr>
        <w:tab/>
        <w:t xml:space="preserve">Scenario</w:t>
      </w:r>
    </w:p>
    <xsl:apply-templates select="@pending, node()">
      <xsl:with-param name="ScenarioLevel" tunnel="yes" select="$ScenarioLevel+1"/>
    </xsl:apply-templates>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="{$style}-end"/>
      </w:pPr>
      <w:r>
        <w:t>
          <xsl:value-of select="@label"/>
        </w:t>
      </w:r>
      <w:r>
        <w:rPr>
          <w:rStyle w:val="x:container-label"/>
        </w:rPr>
        <w:tab/>
        <w:t xml:space="preserve">End Scenario</w:t>
      </w:r>
    </w:p>
  </xsl:template>

  <xsl:template match="x:scenario/@pending">
    <w:p>
      <w:pPr>
        <w:pStyle w:val="x:scenario-pending"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:i/>
        </w:rPr>
        <w:t>Pending</w:t>
        <w:tab/>
        <w:t>
          <xsl:value-of select="string()"/>
        </w:t>
      </w:r>
    </w:p>
  </xsl:template>

  <xsl:template match="x:call[@template]">
    <w:p>
      <w:pPr>
        <w:pStyle w:val="x:call-template"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rStyle w:val="x:inline-code"/>
        </w:rPr>
        <w:t>{@template}</w:t>
      </w:r>
      <w:r>
        <w:rPr>
          <w:rStyle w:val="x:container-label"/>
        </w:rPr>
        <w:tab/>
        <w:t>Named Template</w:t>
      </w:r>
    </w:p>
    <xsl:apply-templates/>
  </xsl:template>
  
  <xsl:template match="x:call[@function]">
    <w:p>
      <w:pPr>
        <w:pStyle w:val="x:call-function"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rStyle w:val="x:inline-code"/>
        </w:rPr>
        <w:t>{@function}</w:t>
      </w:r>
      <w:r>
        <w:rPr>
          <w:rStyle w:val="x:container-label"/>
        </w:rPr>
        <w:tab/>
        <w:t>Function</w:t>
      </w:r>
    </w:p>
    <xsl:apply-templates/>
  </xsl:template>
  
  <xsl:template match="x:call/x:param">
    <w:p>
      <w:pPr>
        <w:pStyle w:val="x:call-template"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rStyle w:val="x:inline-code"/>
        </w:rPr>
        <w:t>${@name}</w:t>
      </w:r>
      <w:r>
        <w:t xml:space="preserve"> has value: </w:t>
      </w:r>
      <xsl:apply-templates select="@select"/>
      <w:r>
        <w:tab/>
      </w:r>
      <w:r>
        <w:rPr>
          <w:rStyle w:val="x:inline-code"/>
        </w:rPr>
        <w:t>{@as}</w:t>
      </w:r>
      <xsl:if test="@as">
        <w:r>
          <w:t xml:space="preserve"> type</w:t>
        </w:r>
      </xsl:if>
      <w:r>
        <w:rPr>
          <w:rStyle w:val="x:container-label"/>
        </w:rPr>
        <xsl:if test="@tunnel = ('yes', 'true', 1, '1')">
          <w:t xml:space="preserve"> Tunnel</w:t>
        </xsl:if>
        <w:t xml:space="preserve"> Parameter</w:t>
      </w:r>
    </w:p>
    <xsl:apply-templates select="@href"/>
    <xsl:call-template name="xml"/>
  </xsl:template>
  
  <xsl:template match="x:param/@href">
    <xsl:param name="pStyle" select="'x:call-template'"/>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="{$pStyle}"/>
      </w:pPr>
      <w:r>
        <w:t>  In file: {.}</w:t>
      </w:r>
    </w:p>
  </xsl:template>
  
  <xsl:template match="x:param/@select">
    <w:r>
      <w:rPr>
        <w:rStyle w:val="x:inline-code"/>
      </w:rPr>
      <w:t>{.}</w:t>
    </w:r>
  </xsl:template>

  <xsl:template match="x:context">
    <w:p>
      <w:pPr>
        <w:pStyle w:val="x:context"/>
      </w:pPr>
      <xsl:if test="@mode">
        <w:r>
          <w:t xml:space="preserve">In mode:</w:t>
        </w:r>
      </xsl:if>
      <xsl:apply-templates select="@mode"/>
      <w:r>
        <w:rPr>
          <w:rStyle w:val="x:container-label"/>
        </w:rPr>
        <w:tab/>
        <xsl:if test="x2w:isWord(*)">
          <w:t xml:space="preserve">Word </w:t>
        </xsl:if>
        <w:t>Input</w:t>
      </w:r>
    </w:p>
    <xsl:choose>
      <xsl:when test="x2w:isWord(*)">
        <xsl:apply-templates/>
      </xsl:when>
      <xsl:when test="* | processing-instruction() | comment()">
        <xsl:call-template name="xml"/>
      </xsl:when>
    </xsl:choose>
    <xsl:if test="string-length(normalize-space(.)) gt 0">
      <w:p>
        <w:pPr>
          <w:pStyle w:val="x:context-end"/>
        </w:pPr>
      </w:p>
    </xsl:if>
  </xsl:template>

  <xsl:template match="@mode">
    <w:r>
      <w:t xml:space="preserve"> <xsl:value-of select="."/></w:t>
    </w:r>
  </xsl:template>

  <xsl:template match="x:expect">
    <w:p>
      <w:pPr>
        <w:pStyle w:val="x:expect"/>
      </w:pPr>
      <xsl:apply-templates select="@label"/>
      <w:r>
        <w:rPr>
          <w:rStyle w:val="x:container-label"/>
        </w:rPr>
        <w:tab/>
        <w:t xml:space="preserve">Expected </w:t>
        <xsl:if test="x2w:isWord(*)">
          <w:t xml:space="preserve">Word </w:t>
        </xsl:if>
        <w:t>Output</w:t>
      </w:r>
    </w:p>
    <xsl:apply-templates select="@select, @test"/>
    <xsl:choose>
      <xsl:when test="x2w:isWord(*)">
        <xsl:apply-templates/>
      </xsl:when>
      <xsl:when test="* | processing-instruction() | comment()">
        <xsl:call-template name="xml"/>
      </xsl:when>
    </xsl:choose>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="x:expect-end"/>
      </w:pPr>
    </w:p>
  </xsl:template>

  <xsl:template match="x:expect/@label">
    <w:r>
      <w:t>
        <xsl:value-of select="."/>
      </w:t>
    </w:r>
  </xsl:template>

  <xsl:template match="x:expect/@select | x:expect/@test">
    <w:p>
      <w:pPr>
        <w:pStyle w:val="x:expect-details"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rStyle w:val="x:inline-code"/>
        </w:rPr>
        <w:t>
          <xsl:text>@</xsl:text>
          <xsl:value-of select="name()"/>
          <xsl:text> = </xsl:text>
          <xsl:value-of select="string(.)"/>
        </w:t>
      </w:r>
    </w:p>
  </xsl:template>

  <xsl:template match="x:pending">
    <w:p>
      <w:pPr>
        <w:pStyle w:val="x:pending"/>
      </w:pPr>
      <w:r>
        <w:t>
          <xsl:value-of select="@label"/>
        </w:t>
      </w:r>
      <w:r>
        <w:rPr>
          <w:i/>
        </w:rPr>
        <w:tab/>
        <w:t xml:space="preserve">Pending</w:t>
      </w:r>
    </w:p>
    <xsl:apply-templates select="node()"/>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="x:pending-end"/>
      </w:pPr>
    </w:p>
  </xsl:template>

  <xsl:template name="xml">
    <xsl:if test="*">
      <w:p>
        <w:pPr>
          <w:pStyle w:val="x:code"/>
        </w:pPr>
        <xsl:apply-templates select="node()" mode="xml"/>
      </w:p>
    </xsl:if>
  </xsl:template>

  <xsl:template match="*" mode="xml">
    <w:r>
      <w:rPr>
        <w:rStyle w:val="x:code-element-furniture"/>
      </w:rPr>
      <w:t>&lt;</w:t>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rStyle w:val="x:code-element-name"/>
      </w:rPr>
      <w:t>
        <xsl:value-of select="name()"/>
      </w:t>
    </w:r>
    <xsl:apply-templates select="@*" mode="xml"/>
    <xsl:choose>
      <xsl:when test="child::node()">
        <w:r>
          <w:rPr>
            <w:rStyle w:val="x:code-element-furniture"/>
          </w:rPr>
          <w:t>&gt;</w:t>
        </w:r>
        <xsl:apply-templates select="node()" mode="xml"/>
        <w:r>
          <w:rPr>
            <w:rStyle w:val="x:code-element-furniture"/>
          </w:rPr>
          <w:t>&lt;/</w:t>
        </w:r>
        <w:r>
          <w:rPr>
            <w:rStyle w:val="x:code-element-name"/>
          </w:rPr>
          <w:t>
            <xsl:value-of select="name()"/>
          </w:t>
        </w:r>
        <w:r>
          <w:rPr>
            <w:rStyle w:val="x:code-element-furniture"/>
          </w:rPr>
          <w:t>&gt;</w:t>
        </w:r>
      </xsl:when>
      <xsl:otherwise>
        <w:r>
          <w:rPr>
            <w:rStyle w:val="x:code-element-furniture"/>
          </w:rPr>
          <w:t>/&gt;</w:t>
        </w:r>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="@*" mode="xml">
    <w:r>
      <w:rPr>
        <w:rStyle w:val="x:code-attribute-name"/>
      </w:rPr>
      <w:t xml:space="preserve"> <xsl:value-of select="name()"/></w:t>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rStyle w:val="x:code-attribute-furniture"/>
      </w:rPr>
      <w:t xml:space="preserve">="</w:t>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rStyle w:val="x:code-attribute-content"/>
      </w:rPr>
      <w:t xml:space="preserve"><xsl:value-of select="string()"/></w:t>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rStyle w:val="x:code-attribute-furniture"/>
      </w:rPr>
      <w:t>"</w:t>
    </w:r>
  </xsl:template>

  <xsl:template match="text()" mode="xml">
    <w:r>
      <xsl:analyze-string select="." regex="[&#10;]+">
        <xsl:matching-substring>
          <w:br/>
        </xsl:matching-substring>
        <xsl:non-matching-substring>
          <w:t xml:space="preserve"><xsl:copy/></w:t>
        </xsl:non-matching-substring>
      </xsl:analyze-string>
    </w:r>
  </xsl:template>

  <xsl:template match="comment()" mode="xml">
    <w:r>
      <w:rPr>
        <w:rStyle w:val="x:code-comment"/>
      </w:rPr>
      <w:t xml:space="preserve">&lt;!--<xsl:value-of select="."/>--!&gt;</w:t>
    </w:r>
  </xsl:template>

  <xsl:template match="processing-instruction()">
    <w:r>
      <w:rPr>
        <w:rStyle w:val="x:code-pi"/>
      </w:rPr>
      <w:t>&lt;?<xsl:value-of select="name()"/>
        <xsl:value-of select="."/>?&gt;</w:t>
    </w:r>
  </xsl:template>

  <xsl:function name="x2w:isWord" as="xs:boolean">
    <xsl:param name="nodes" as="node()*"/>
    <xsl:choose>
      <xsl:when test="$nodes/namespace-uri() = (
          'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
          'http://schemas.microsoft.com/office/word/2003/wordml'
        )">
        <xsl:value-of select="true()"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="false()"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

</xsl:stylesheet>
