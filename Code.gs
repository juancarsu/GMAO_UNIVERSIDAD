{\rtf1\ansi\ansicpg1252\cocoartf2867
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fnil\fcharset0 Menlo-Regular;}
{\colortbl;\red255\green255\blue255;\red77\green80\blue85;\red246\green247\blue249;\red46\green49\blue51;
\red20\green67\blue174;\red186\green6\blue115;\red24\green25\blue27;\red162\green0\blue16;\red18\green115\blue126;
}
{\*\expandedcolortbl;;\cssrgb\c37255\c38824\c40784;\cssrgb\c97255\c97647\c98039;\cssrgb\c23529\c25098\c26275;
\cssrgb\c9412\c35294\c73725;\cssrgb\c78824\c15294\c52549;\cssrgb\c12549\c12941\c14118;\cssrgb\c70196\c7843\c7059;\cssrgb\c3529\c52157\c56863;
}
\paperw11900\paperh16840\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\deftab720
\pard\pardeftab720\partightenfactor0

\f0\fs26 \cf2 \cb3 \expnd0\expndtw0\kerning0
\outl0\strokewidth0 \strokec2 // ==========================================\cf4 \cb1 \strokec4 \
\cf2 \cb3 \strokec2 // 1. CONFIGURACI\'d3N Y ROUTING\cf4 \cb1 \strokec4 \
\cf2 \cb3 \strokec2 // ==========================================\cf4 \cb1 \strokec4 \
\pard\pardeftab720\partightenfactor0
\cf5 \cb3 \strokec5 const\cf4 \strokec4  \cf6 \strokec6 PROPS\cf4 \strokec4  = \cf6 \strokec6 PropertiesService\cf4 \strokec4 .\cf7 \strokec7 getScriptProperties\cf4 \strokec4 ();\cb1 \
\
\cf5 \cb3 \strokec5 function\cf4 \strokec4  \cf7 \strokec7 doGet\cf4 \strokec4 (\cf7 \strokec7 e\cf4 \strokec4 ) \{\cb1 \
\pard\pardeftab720\partightenfactor0
\cf4 \cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 dbId\cf4 \strokec4  = \cf6 \strokec6 PROPS\cf4 \strokec4 .\cf7 \strokec7 getProperty\cf4 \strokec4 (\cf8 \strokec8 'DB_SS_ID'\cf4 \strokec4 );\cb1 \
\cb3   \cf5 \strokec5 if\cf4 \strokec4  (!\cf7 \strokec7 dbId\cf4 \strokec4 ) \{\cb1 \
\cb3     \cf5 \strokec5 return\cf4 \strokec4  \cf6 \strokec6 HtmlService\cf4 \strokec4 .\cf7 \strokec7 createTemplateFromFile\cf4 \strokec4 (\cf8 \strokec8 'Setup'\cf4 \strokec4 )\cb1 \
\cb3       .\cf7 \strokec7 evaluate\cf4 \strokec4 ().\cf7 \strokec7 setTitle\cf4 \strokec4 (\cf8 \strokec8 'Instalaci\'f3n GMAO'\cf4 \strokec4 ).\cf7 \strokec7 setXFrameOptionsMode\cf4 \strokec4 (\cf6 \strokec6 HtmlService\cf4 \strokec4 .\cf6 \strokec6 XFrameOptionsMode\cf4 \strokec4 .\cf6 \strokec6 ALLOWALL\cf4 \strokec4 );\cb1 \
\cb3   \}\cb1 \
\cb3   \cf5 \strokec5 return\cf4 \strokec4  \cf6 \strokec6 HtmlService\cf4 \strokec4 .\cf7 \strokec7 createTemplateFromFile\cf4 \strokec4 (\cf8 \strokec8 'Index'\cf4 \strokec4 )\cb1 \
\cb3       .\cf7 \strokec7 evaluate\cf4 \strokec4 ().\cf7 \strokec7 setTitle\cf4 \strokec4 (\cf8 \strokec8 'GMAO Universidad'\cf4 \strokec4 ).\cf7 \strokec7 addMetaTag\cf4 \strokec4 (\cf8 \strokec8 'viewport'\cf4 \strokec4 , \cf8 \strokec8 'width=device-width, initial-scale=1'\cf4 \strokec4 )\cb1 \
\cb3       .\cf7 \strokec7 setXFrameOptionsMode\cf4 \strokec4 (\cf6 \strokec6 HtmlService\cf4 \strokec4 .\cf6 \strokec6 XFrameOptionsMode\cf4 \strokec4 .\cf6 \strokec6 ALLOWALL\cf4 \strokec4 );\cb1 \
\cb3 \}\cb1 \
\
\pard\pardeftab720\partightenfactor0
\cf2 \cb3 \strokec2 // *** NUEVA FUNCI\'d3N IMPORTANTE ***\cf4 \cb1 \strokec4 \
\pard\pardeftab720\partightenfactor0
\cf5 \cb3 \strokec5 function\cf4 \strokec4  \cf7 \strokec7 include\cf4 \strokec4 (\cf7 \strokec7 filename\cf4 \strokec4 ) \{\cb1 \
\pard\pardeftab720\partightenfactor0
\cf4 \cb3   \cf5 \strokec5 return\cf4 \strokec4  \cf6 \strokec6 HtmlService\cf4 \strokec4 .\cf7 \strokec7 createHtmlOutputFromFile\cf4 \strokec4 (\cf7 \strokec7 filename\cf4 \strokec4 ).\cf7 \strokec7 getContent\cf4 \strokec4 ();\cb1 \
\cb3 \}\cb1 \
\
\pard\pardeftab720\partightenfactor0
\cf2 \cb3 \strokec2 // ==========================================\cf4 \cb1 \strokec4 \
\cf2 \cb3 \strokec2 // 2. INSTALACI\'d3N (SETUP)\cf4 \cb1 \strokec4 \
\cf2 \cb3 \strokec2 // ==========================================\cf4 \cb1 \strokec4 \
\pard\pardeftab720\partightenfactor0
\cf5 \cb3 \strokec5 function\cf4 \strokec4  \cf7 \strokec7 instalarSistema\cf4 \strokec4 (\cf7 \strokec7 urlSheet\cf4 \strokec4 ) \{\cb1 \
\pard\pardeftab720\partightenfactor0
\cf4 \cb3   \cf5 \strokec5 try\cf4 \strokec4  \{\cb1 \
\cb3     \cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 ss\cf4 \strokec4 ;\cb1 \
\cb3     \cf5 \strokec5 try\cf4 \strokec4  \{ \cf7 \strokec7 ss\cf4 \strokec4  = \cf6 \strokec6 SpreadsheetApp\cf4 \strokec4 .\cf7 \strokec7 openByUrl\cf4 \strokec4 (\cf7 \strokec7 urlSheet\cf4 \strokec4 ); \} \cb1 \
\cb3     \cf5 \strokec5 catch\cf4 \strokec4 (\cf7 \strokec7 e\cf4 \strokec4 ) \{ \cf7 \strokec7 ss\cf4 \strokec4  = \cf6 \strokec6 SpreadsheetApp\cf4 \strokec4 .\cf7 \strokec7 openById\cf4 \strokec4 (\cf7 \strokec7 urlSheet\cf4 \strokec4 ); \}\cb1 \
\cb3     \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 ssId\cf4 \strokec4  = \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 getId\cf4 \strokec4 ();\cb1 \
\cb3     \cb1 \
\cb3     \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 tabs\cf4 \strokec4  = \{\cb1 \
\cb3       \cf8 \strokec8 'CONFIG'\cf4 \strokec4 : [\cf8 \strokec8 'Clave'\cf4 \strokec4 , \cf8 \strokec8 'Valor'\cf4 \strokec4 , \cf8 \strokec8 'Descripcion'\cf4 \strokec4 ],\cb1 \
\cb3       \cf8 \strokec8 'CAMPUS'\cf4 \strokec4 : [\cf8 \strokec8 'ID'\cf4 \strokec4 , \cf8 \strokec8 'Nombre'\cf4 \strokec4 , \cf8 \strokec8 'Provincia'\cf4 \strokec4 , \cf8 \strokec8 'Direccion'\cf4 \strokec4 , \cf8 \strokec8 'ID_Carpeta_Drive'\cf4 \strokec4 ],\cb1 \
\cb3       \cf8 \strokec8 'EDIFICIOS'\cf4 \strokec4 : [\cf8 \strokec8 'ID'\cf4 \strokec4 , \cf8 \strokec8 'ID_Campus'\cf4 \strokec4 , \cf8 \strokec8 'Nombre'\cf4 \strokec4 , \cf8 \strokec8 'Contacto'\cf4 \strokec4 , \cf8 \strokec8 'ID_Carpeta_Drive'\cf4 \strokec4 , \cf8 \strokec8 'ID_Carpeta_Activos'\cf4 \strokec4 ],\cb1 \
\cb3       \cf8 \strokec8 'CAT_INSTALACIONES'\cf4 \strokec4 : [\cf8 \strokec8 'ID'\cf4 \strokec4 , \cf8 \strokec8 'Nombre'\cf4 \strokec4 , \cf8 \strokec8 'Normativa_Ref'\cf4 \strokec4 , \cf8 \strokec8 'Periodicidad_Dias'\cf4 \strokec4 ],\cb1 \
\cb3       \cf8 \strokec8 'ACTIVOS'\cf4 \strokec4 : [\cf8 \strokec8 'ID'\cf4 \strokec4 , \cf8 \strokec8 'ID_Edificio'\cf4 \strokec4 , \cf8 \strokec8 'Tipo'\cf4 \strokec4 , \cf8 \strokec8 'Nombre'\cf4 \strokec4 , \cf8 \strokec8 'Marca'\cf4 \strokec4 , \cf8 \strokec8 'Fecha_Alta'\cf4 \strokec4 , \cf8 \strokec8 'ID_Carpeta_Drive'\cf4 \strokec4 ],\cb1 \
\cb3       \cf8 \strokec8 'DOCS_HISTORICO'\cf4 \strokec4 : [\cf8 \strokec8 'ID_Doc'\cf4 \strokec4 , \cf8 \strokec8 'Tipo_Entidad'\cf4 \strokec4 , \cf8 \strokec8 'ID_Entidad'\cf4 \strokec4 , \cf8 \strokec8 'Nombre_Archivo'\cf4 \strokec4 , \cf8 \strokec8 'URL'\cf4 \strokec4 , \cf8 \strokec8 'Version'\cf4 \strokec4 , \cf8 \strokec8 'Fecha'\cf4 \strokec4 , \cf8 \strokec8 'Usuario'\cf4 \strokec4 , \cf8 \strokec8 'ID_Archivo_Drive'\cf4 \strokec4 ],\cb1 \
\cb3       \cf8 \strokec8 'PLAN_MANTENIMIENTO'\cf4 \strokec4 : [\cf8 \strokec8 'ID_Plan'\cf4 \strokec4 , \cf8 \strokec8 'ID_Activo'\cf4 \strokec4 , \cf8 \strokec8 'Tipo_Revision'\cf4 \strokec4 , \cf8 \strokec8 'Fecha_Ultima'\cf4 \strokec4 , \cf8 \strokec8 'Fecha_Proxima'\cf4 \strokec4 , \cf8 \strokec8 'Periodicidad_Dias'\cf4 \strokec4 , \cf8 \strokec8 'Estado'\cf4 \strokec4 ],\cb1 \
\cb3       \cf8 \strokec8 'CONTRATOS'\cf4 \strokec4 : [\cf8 \strokec8 'ID_Contrato'\cf4 \strokec4 , \cf8 \strokec8 'Tipo_Entidad'\cf4 \strokec4 , \cf8 \strokec8 'ID_Entidad'\cf4 \strokec4 , \cf8 \strokec8 'Proveedor'\cf4 \strokec4 , \cf8 \strokec8 'Num_Ref'\cf4 \strokec4 , \cf8 \strokec8 'Fecha_Inicio'\cf4 \strokec4 , \cf8 \strokec8 'Fecha_Fin'\cf4 \strokec4 ]\cb1 \
\cb3     \};\cb1 \
\
\cb3     \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 sheets\cf4 \strokec4  = \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 getSheets\cf4 \strokec4 ().\cf7 \strokec7 map\cf4 \strokec4 (\cf7 \strokec7 s\cf4 \strokec4  => \cf7 \strokec7 s\cf4 \strokec4 .\cf7 \strokec7 getName\cf4 \strokec4 ());\cb1 \
\cb3     \cf5 \strokec5 for\cf4 \strokec4  (\cf5 \strokec5 const\cf4 \strokec4  [\cf7 \strokec7 name\cf4 \strokec4 , \cf7 \strokec7 headers\cf4 \strokec4 ] \cf5 \strokec5 of\cf4 \strokec4  \cf6 \strokec6 Object\cf4 \strokec4 .\cf7 \strokec7 entries\cf4 \strokec4 (\cf7 \strokec7 tabs\cf4 \strokec4 )) \{\cb1 \
\cb3       \cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 sheet\cf4 \strokec4  = \cf7 \strokec7 sheets\cf4 \strokec4 .\cf7 \strokec7 includes\cf4 \strokec4 (\cf7 \strokec7 name\cf4 \strokec4 ) ? \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 getSheetByName\cf4 \strokec4 (\cf7 \strokec7 name\cf4 \strokec4 ) : \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 insertSheet\cf4 \strokec4 (\cf7 \strokec7 name\cf4 \strokec4 );\cb1 \
\cb3       \cf5 \strokec5 if\cf4 \strokec4  (\cf7 \strokec7 sheet\cf4 \strokec4 .\cf7 \strokec7 getLastRow\cf4 \strokec4 () === \cf9 \strokec9 0\cf4 \strokec4 ) \{\cb1 \
\cb3         \cf7 \strokec7 sheet\cf4 \strokec4 .\cf7 \strokec7 getRange\cf4 \strokec4 (\cf9 \strokec9 1\cf4 \strokec4 , \cf9 \strokec9 1\cf4 \strokec4 , \cf9 \strokec9 1\cf4 \strokec4 , \cf7 \strokec7 headers\cf4 \strokec4 .\cf7 \strokec7 length\cf4 \strokec4 ).\cf7 \strokec7 setValues\cf4 \strokec4 ([\cf7 \strokec7 headers\cf4 \strokec4 ]).\cf7 \strokec7 setFontWeight\cf4 \strokec4 (\cf8 \strokec8 'bold'\cf4 \strokec4 ).\cf7 \strokec7 setBackground\cf4 \strokec4 (\cf8 \strokec8 '#ddd'\cf4 \strokec4 );\cb1 \
\cb3       \}\cb1 \
\cb3     \}\cb1 \
\
\cb3     \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 sheetConfig\cf4 \strokec4  = \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 getSheetByName\cf4 \strokec4 (\cf8 \strokec8 'CONFIG'\cf4 \strokec4 );\cb1 \
\cb3     \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 dataConfig\cf4 \strokec4  = \cf7 \strokec7 sheetConfig\cf4 \strokec4 .\cf7 \strokec7 getDataRange\cf4 \strokec4 ().\cf7 \strokec7 getValues\cf4 \strokec4 ();\cb1 \
\cb3     \cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 rootId\cf4 \strokec4  = \cf5 \strokec5 null\cf4 \strokec4 ;\cb1 \
\cb3     \cf5 \strokec5 for\cf4 \strokec4 (\cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 row\cf4 \strokec4  \cf5 \strokec5 of\cf4 \strokec4  \cf7 \strokec7 dataConfig\cf4 \strokec4 ) \{ \cf5 \strokec5 if\cf4 \strokec4 (\cf7 \strokec7 row\cf4 \strokec4 [\cf9 \strokec9 0\cf4 \strokec4 ] === \cf8 \strokec8 'ROOT_FOLDER_ID'\cf4 \strokec4 ) \cf7 \strokec7 rootId\cf4 \strokec4  = \cf7 \strokec7 row\cf4 \strokec4 [\cf9 \strokec9 1\cf4 \strokec4 ]; \}\cb1 \
\
\cb3     \cf5 \strokec5 if\cf4 \strokec4 (!\cf7 \strokec7 rootId\cf4 \strokec4 ) \{\cb1 \
\cb3       \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 rootFolder\cf4 \strokec4  = \cf6 \strokec6 DriveApp\cf4 \strokec4 .\cf7 \strokec7 createFolder\cf4 \strokec4 (\cf8 \strokec8 "GMAO UNIVERSIDAD - GESTI\'d3N DOCUMENTAL"\cf4 \strokec4 );\cb1 \
\cb3       \cf7 \strokec7 rootId\cf4 \strokec4  = \cf7 \strokec7 rootFolder\cf4 \strokec4 .\cf7 \strokec7 getId\cf4 \strokec4 ();\cb1 \
\cb3       \cf7 \strokec7 sheetConfig\cf4 \strokec4 .\cf7 \strokec7 appendRow\cf4 \strokec4 ([\cf8 \strokec8 'ROOT_FOLDER_ID'\cf4 \strokec4 , \cf7 \strokec7 rootId\cf4 \strokec4 , \cf8 \strokec8 'Ra\'edz del sistema'\cf4 \strokec4 ]);\cb1 \
\cb3     \}\cb1 \
\cb3     \cb1 \
\cb3     \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 emailAdmin\cf4 \strokec4  = \cf6 \strokec6 Session\cf4 \strokec4 .\cf7 \strokec7 getActiveUser\cf4 \strokec4 ().\cf7 \strokec7 getEmail\cf4 \strokec4 ();\cb1 \
\cb3     \cf7 \strokec7 sheetConfig\cf4 \strokec4 .\cf7 \strokec7 appendRow\cf4 \strokec4 ([\cf8 \strokec8 'EMAIL_AVISOS'\cf4 \strokec4 , \cf7 \strokec7 emailAdmin\cf4 \strokec4 , \cf8 \strokec8 'Email para alertas semanales'\cf4 \strokec4 ]);\cb1 \
\
\cb3     \cf6 \strokec6 PROPS\cf4 \strokec4 .\cf7 \strokec7 setProperty\cf4 \strokec4 (\cf8 \strokec8 'DB_SS_ID'\cf4 \strokec4 , \cf7 \strokec7 ssId\cf4 \strokec4 );\cb1 \
\cb3     \cf5 \strokec5 return\cf4 \strokec4  \{ \cf7 \strokec7 success\cf4 \strokec4 : \cf5 \strokec5 true\cf4 \strokec4  \};\cb1 \
\cb3   \} \cf5 \strokec5 catch\cf4 \strokec4  (\cf7 \strokec7 e\cf4 \strokec4 ) \{\cb1 \
\cb3     \cf5 \strokec5 return\cf4 \strokec4  \{ \cf7 \strokec7 success\cf4 \strokec4 : \cf5 \strokec5 false\cf4 \strokec4 , \cf7 \strokec7 error\cf4 \strokec4 : \cf7 \strokec7 e\cf4 \strokec4 .\cf7 \strokec7 toString\cf4 \strokec4 () \};\cb1 \
\cb3   \}\cb1 \
\cb3 \}\cb1 \
\
\pard\pardeftab720\partightenfactor0
\cf2 \cb3 \strokec2 // ==========================================\cf4 \cb1 \strokec4 \
\cf2 \cb3 \strokec2 // 3. L\'d3GICA DE DRIVE\cf4 \cb1 \strokec4 \
\cf2 \cb3 \strokec2 // ==========================================\cf4 \cb1 \strokec4 \
\pard\pardeftab720\partightenfactor0
\cf5 \cb3 \strokec5 function\cf4 \strokec4  \cf7 \strokec7 getRootFolderId\cf4 \strokec4 () \{\cb1 \
\pard\pardeftab720\partightenfactor0
\cf4 \cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 ss\cf4 \strokec4  = \cf6 \strokec6 SpreadsheetApp\cf4 \strokec4 .\cf7 \strokec7 openById\cf4 \strokec4 (\cf6 \strokec6 PROPS\cf4 \strokec4 .\cf7 \strokec7 getProperty\cf4 \strokec4 (\cf8 \strokec8 'DB_SS_ID'\cf4 \strokec4 ));\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 data\cf4 \strokec4  = \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 getSheetByName\cf4 \strokec4 (\cf8 \strokec8 'CONFIG'\cf4 \strokec4 ).\cf7 \strokec7 getDataRange\cf4 \strokec4 ().\cf7 \strokec7 getValues\cf4 \strokec4 ();\cb1 \
\cb3   \cf5 \strokec5 for\cf4 \strokec4 (\cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 row\cf4 \strokec4  \cf5 \strokec5 of\cf4 \strokec4  \cf7 \strokec7 data\cf4 \strokec4 ) \{ \cf5 \strokec5 if\cf4 \strokec4 (\cf7 \strokec7 row\cf4 \strokec4 [\cf9 \strokec9 0\cf4 \strokec4 ] === \cf8 \strokec8 'ROOT_FOLDER_ID'\cf4 \strokec4 ) \cf5 \strokec5 return\cf4 \strokec4  \cf7 \strokec7 row\cf4 \strokec4 [\cf9 \strokec9 1\cf4 \strokec4 ]; \}\cb1 \
\cb3   \cf5 \strokec5 return\cf4 \strokec4  \cf5 \strokec5 null\cf4 \strokec4 ;\cb1 \
\cb3 \}\cb1 \
\
\pard\pardeftab720\partightenfactor0
\cf5 \cb3 \strokec5 function\cf4 \strokec4  \cf7 \strokec7 crearCarpeta\cf4 \strokec4 (\cf7 \strokec7 nombre\cf4 \strokec4 , \cf7 \strokec7 idPadre\cf4 \strokec4 ) \{\cb1 \
\pard\pardeftab720\partightenfactor0
\cf4 \cb3   \cf5 \strokec5 return\cf4 \strokec4  \cf6 \strokec6 DriveApp\cf4 \strokec4 .\cf7 \strokec7 getFolderById\cf4 \strokec4 (\cf7 \strokec7 idPadre\cf4 \strokec4 ).\cf7 \strokec7 createFolder\cf4 \strokec4 (\cf7 \strokec7 nombre\cf4 \strokec4 ).\cf7 \strokec7 getId\cf4 \strokec4 ();\cb1 \
\cb3 \}\cb1 \
\
\pard\pardeftab720\partightenfactor0
\cf2 \cb3 \strokec2 // ==========================================\cf4 \cb1 \strokec4 \
\cf2 \cb3 \strokec2 // 4. CAT\'c1LOGO Y NORMATIVA\cf4 \cb1 \strokec4 \
\cf2 \cb3 \strokec2 // ==========================================\cf4 \cb1 \strokec4 \
\pard\pardeftab720\partightenfactor0
\cf5 \cb3 \strokec5 function\cf4 \strokec4  \cf7 \strokec7 getCatalogoInstalaciones\cf4 \strokec4 () \{\cb1 \
\pard\pardeftab720\partightenfactor0
\cf4 \cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 ss\cf4 \strokec4  = \cf6 \strokec6 SpreadsheetApp\cf4 \strokec4 .\cf7 \strokec7 openById\cf4 \strokec4 (\cf6 \strokec6 PROPS\cf4 \strokec4 .\cf7 \strokec7 getProperty\cf4 \strokec4 (\cf8 \strokec8 'DB_SS_ID'\cf4 \strokec4 ));\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 sheet\cf4 \strokec4  = \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 getSheetByName\cf4 \strokec4 (\cf8 \strokec8 'CAT_INSTALACIONES'\cf4 \strokec4 );\cb1 \
\cb3   \cb1 \
\cb3   \cf5 \strokec5 if\cf4 \strokec4  (\cf7 \strokec7 sheet\cf4 \strokec4 .\cf7 \strokec7 getLastRow\cf4 \strokec4 () < \cf9 \strokec9 2\cf4 \strokec4 ) \{\cb1 \
\cb3     \cf7 \strokec7 sheet\cf4 \strokec4 .\cf7 \strokec7 appendRow\cf4 \strokec4 ([\cf6 \strokec6 Utilities\cf4 \strokec4 .\cf7 \strokec7 getUuid\cf4 \strokec4 (), \cf8 \strokec8 'Baja Tensi\'f3n (REBT)'\cf4 \strokec4 , \cf8 \strokec8 'RD 842/2002'\cf4 \strokec4 , \cf9 \strokec9 1825\cf4 \strokec4 ]);\cb1 \
\cb3     \cf7 \strokec7 sheet\cf4 \strokec4 .\cf7 \strokec7 appendRow\cf4 \strokec4 ([\cf6 \strokec6 Utilities\cf4 \strokec4 .\cf7 \strokec7 getUuid\cf4 \strokec4 (), \cf8 \strokec8 'Ascensores'\cf4 \strokec4 , \cf8 \strokec8 'ITC AEM 1'\cf4 \strokec4 , \cf9 \strokec9 30\cf4 \strokec4 ]);\cb1 \
\cb3     \cf7 \strokec7 sheet\cf4 \strokec4 .\cf7 \strokec7 appendRow\cf4 \strokec4 ([\cf6 \strokec6 Utilities\cf4 \strokec4 .\cf7 \strokec7 getUuid\cf4 \strokec4 (), \cf8 \strokec8 'Instalaciones T\'e9rmicas (RITE)'\cf4 \strokec4 , \cf8 \strokec8 'RD 1027/2007'\cf4 \strokec4 , \cf9 \strokec9 365\cf4 \strokec4 ]);\cb1 \
\cb3     \cf7 \strokec7 sheet\cf4 \strokec4 .\cf7 \strokec7 appendRow\cf4 \strokec4 ([\cf6 \strokec6 Utilities\cf4 \strokec4 .\cf7 \strokec7 getUuid\cf4 \strokec4 (), \cf8 \strokec8 'Protecci\'f3n Incendios (PCI)'\cf4 \strokec4 , \cf8 \strokec8 'RIPCI'\cf4 \strokec4 , \cf9 \strokec9 90\cf4 \strokec4 ]);\cb1 \
\cb3   \}\cb1 \
\cb3   \cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 data\cf4 \strokec4  = \cf7 \strokec7 sheet\cf4 \strokec4 .\cf7 \strokec7 getDataRange\cf4 \strokec4 ().\cf7 \strokec7 getValues\cf4 \strokec4 ();\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 catalogo\cf4 \strokec4  = [];\cb1 \
\cb3   \cf5 \strokec5 for\cf4 \strokec4 (\cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 i\cf4 \strokec4 =\cf9 \strokec9 1\cf4 \strokec4 ; \cf7 \strokec7 i\cf4 \strokec4 <\cf7 \strokec7 data\cf4 \strokec4 .\cf7 \strokec7 length\cf4 \strokec4 ; \cf7 \strokec7 i\cf4 \strokec4 ++)\{\cb1 \
\cb3     \cf7 \strokec7 catalogo\cf4 \strokec4 .\cf7 \strokec7 push\cf4 \strokec4 (\{ \cf7 \strokec7 id\cf4 \strokec4 : \cf7 \strokec7 data\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 0\cf4 \strokec4 ], \cf7 \strokec7 nombre\cf4 \strokec4 : \cf7 \strokec7 data\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 1\cf4 \strokec4 ], \cf7 \strokec7 norma\cf4 \strokec4 : \cf7 \strokec7 data\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 2\cf4 \strokec4 ], \cf7 \strokec7 dias\cf4 \strokec4 : \cf7 \strokec7 data\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 3\cf4 \strokec4 ] \});\cb1 \
\cb3   \}\cb1 \
\cb3   \cf5 \strokec5 return\cf4 \strokec4  \cf7 \strokec7 catalogo\cf4 \strokec4 ;\cb1 \
\cb3 \}\cb1 \
\
\pard\pardeftab720\partightenfactor0
\cf2 \cb3 \strokec2 // ==========================================\cf4 \cb1 \strokec4 \
\cf2 \cb3 \strokec2 // 5. CREACI\'d3N DE ENTIDADES\cf4 \cb1 \strokec4 \
\cf2 \cb3 \strokec2 // ==========================================\cf4 \cb1 \strokec4 \
\pard\pardeftab720\partightenfactor0
\cf5 \cb3 \strokec5 function\cf4 \strokec4  \cf7 \strokec7 crearCampus\cf4 \strokec4 (\cf7 \strokec7 datos\cf4 \strokec4 ) \{ \cb1 \
\pard\pardeftab720\partightenfactor0
\cf4 \cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 rootId\cf4 \strokec4  = \cf7 \strokec7 getRootFolderId\cf4 \strokec4 ();\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 folderId\cf4 \strokec4  = \cf7 \strokec7 crearCarpeta\cf4 \strokec4 (\cf7 \strokec7 datos\cf4 \strokec4 .\cf7 \strokec7 nombre\cf4 \strokec4 , \cf7 \strokec7 rootId\cf4 \strokec4 );\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 id\cf4 \strokec4  = \cf6 \strokec6 Utilities\cf4 \strokec4 .\cf7 \strokec7 getUuid\cf4 \strokec4 ();\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 ss\cf4 \strokec4  = \cf6 \strokec6 SpreadsheetApp\cf4 \strokec4 .\cf7 \strokec7 openById\cf4 \strokec4 (\cf6 \strokec6 PROPS\cf4 \strokec4 .\cf7 \strokec7 getProperty\cf4 \strokec4 (\cf8 \strokec8 'DB_SS_ID'\cf4 \strokec4 ));\cb1 \
\cb3   \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 getSheetByName\cf4 \strokec4 (\cf8 \strokec8 'CAMPUS'\cf4 \strokec4 ).\cf7 \strokec7 appendRow\cf4 \strokec4 ([\cf7 \strokec7 id\cf4 \strokec4 , \cf7 \strokec7 datos\cf4 \strokec4 .\cf7 \strokec7 nombre\cf4 \strokec4 , \cf7 \strokec7 datos\cf4 \strokec4 .\cf7 \strokec7 provincia\cf4 \strokec4 , \cf7 \strokec7 datos\cf4 \strokec4 .\cf7 \strokec7 direccion\cf4 \strokec4 , \cf7 \strokec7 folderId\cf4 \strokec4 ]);\cb1 \
\cb3   \cf5 \strokec5 return\cf4 \strokec4  \{ \cf7 \strokec7 success\cf4 \strokec4 : \cf5 \strokec5 true\cf4 \strokec4  \};\cb1 \
\cb3 \}\cb1 \
\
\pard\pardeftab720\partightenfactor0
\cf5 \cb3 \strokec5 function\cf4 \strokec4  \cf7 \strokec7 crearEdificio\cf4 \strokec4 (\cf7 \strokec7 datos\cf4 \strokec4 ) \{ \cb1 \
\pard\pardeftab720\partightenfactor0
\cf4 \cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 ss\cf4 \strokec4  = \cf6 \strokec6 SpreadsheetApp\cf4 \strokec4 .\cf7 \strokec7 openById\cf4 \strokec4 (\cf6 \strokec6 PROPS\cf4 \strokec4 .\cf7 \strokec7 getProperty\cf4 \strokec4 (\cf8 \strokec8 'DB_SS_ID'\cf4 \strokec4 ));\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 campData\cf4 \strokec4  = \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 getSheetByName\cf4 \strokec4 (\cf8 \strokec8 'CAMPUS'\cf4 \strokec4 ).\cf7 \strokec7 getDataRange\cf4 \strokec4 ().\cf7 \strokec7 getValues\cf4 \strokec4 ();\cb1 \
\cb3   \cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 parentFolderId\cf4 \strokec4  = \cf5 \strokec5 null\cf4 \strokec4 ;\cb1 \
\cb3   \cf5 \strokec5 for\cf4 \strokec4 (\cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 i\cf4 \strokec4 =\cf9 \strokec9 1\cf4 \strokec4 ; \cf7 \strokec7 i\cf4 \strokec4 <\cf7 \strokec7 campData\cf4 \strokec4 .\cf7 \strokec7 length\cf4 \strokec4 ; \cf7 \strokec7 i\cf4 \strokec4 ++)\{\cb1 \
\cb3     \cf5 \strokec5 if\cf4 \strokec4 (\cf6 \strokec6 String\cf4 \strokec4 (\cf7 \strokec7 campData\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 0\cf4 \strokec4 ]) === \cf6 \strokec6 String\cf4 \strokec4 (\cf7 \strokec7 datos\cf4 \strokec4 .\cf7 \strokec7 idCampus\cf4 \strokec4 )) \{ \cf7 \strokec7 parentFolderId\cf4 \strokec4  = \cf7 \strokec7 campData\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 4\cf4 \strokec4 ]; \cf5 \strokec5 break\cf4 \strokec4 ; \}\cb1 \
\cb3   \}\cb1 \
\cb3   \cf5 \strokec5 if\cf4 \strokec4 (!\cf7 \strokec7 parentFolderId\cf4 \strokec4 ) \cf5 \strokec5 return\cf4 \strokec4  \{ \cf7 \strokec7 success\cf4 \strokec4 : \cf5 \strokec5 false\cf4 \strokec4 , \cf7 \strokec7 error\cf4 \strokec4 : \cf8 \strokec8 "Campus no encontrado"\cf4 \strokec4  \};\cb1 \
\cb3   \cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 edifFolderId\cf4 \strokec4  = \cf7 \strokec7 crearCarpeta\cf4 \strokec4 (\cf7 \strokec7 datos\cf4 \strokec4 .\cf7 \strokec7 nombre\cf4 \strokec4 , \cf7 \strokec7 parentFolderId\cf4 \strokec4 );\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 activosFolderId\cf4 \strokec4  = \cf7 \strokec7 crearCarpeta\cf4 \strokec4 (\cf8 \strokec8 "Activos"\cf4 \strokec4 , \cf7 \strokec7 edifFolderId\cf4 \strokec4 );\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 id\cf4 \strokec4  = \cf6 \strokec6 Utilities\cf4 \strokec4 .\cf7 \strokec7 getUuid\cf4 \strokec4 ();\cb1 \
\cb3   \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 getSheetByName\cf4 \strokec4 (\cf8 \strokec8 'EDIFICIOS'\cf4 \strokec4 ).\cf7 \strokec7 appendRow\cf4 \strokec4 ([\cf7 \strokec7 id\cf4 \strokec4 , \cf7 \strokec7 datos\cf4 \strokec4 .\cf7 \strokec7 idCampus\cf4 \strokec4 , \cf7 \strokec7 datos\cf4 \strokec4 .\cf7 \strokec7 nombre\cf4 \strokec4 , \cf7 \strokec7 datos\cf4 \strokec4 .\cf7 \strokec7 contacto\cf4 \strokec4 , \cf7 \strokec7 edifFolderId\cf4 \strokec4 , \cf7 \strokec7 activosFolderId\cf4 \strokec4 ]);\cb1 \
\cb3   \cf5 \strokec5 return\cf4 \strokec4  \{ \cf7 \strokec7 success\cf4 \strokec4 : \cf5 \strokec5 true\cf4 \strokec4  \};\cb1 \
\cb3 \}\cb1 \
\
\pard\pardeftab720\partightenfactor0
\cf5 \cb3 \strokec5 function\cf4 \strokec4  \cf7 \strokec7 crearActivo\cf4 \strokec4 (\cf7 \strokec7 datos\cf4 \strokec4 ) \{ \cb1 \
\pard\pardeftab720\partightenfactor0
\cf4 \cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 ss\cf4 \strokec4  = \cf6 \strokec6 SpreadsheetApp\cf4 \strokec4 .\cf7 \strokec7 openById\cf4 \strokec4 (\cf6 \strokec6 PROPS\cf4 \strokec4 .\cf7 \strokec7 getProperty\cf4 \strokec4 (\cf8 \strokec8 'DB_SS_ID'\cf4 \strokec4 ));\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 edifData\cf4 \strokec4  = \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 getSheetByName\cf4 \strokec4 (\cf8 \strokec8 'EDIFICIOS'\cf4 \strokec4 ).\cf7 \strokec7 getDataRange\cf4 \strokec4 ().\cf7 \strokec7 getValues\cf4 \strokec4 ();\cb1 \
\cb3   \cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 parentFolderId\cf4 \strokec4  = \cf5 \strokec5 null\cf4 \strokec4 ;\cb1 \
\cb3   \cf5 \strokec5 for\cf4 \strokec4 (\cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 i\cf4 \strokec4 =\cf9 \strokec9 1\cf4 \strokec4 ; \cf7 \strokec7 i\cf4 \strokec4 <\cf7 \strokec7 edifData\cf4 \strokec4 .\cf7 \strokec7 length\cf4 \strokec4 ; \cf7 \strokec7 i\cf4 \strokec4 ++)\{\cb1 \
\cb3     \cf5 \strokec5 if\cf4 \strokec4 (\cf6 \strokec6 String\cf4 \strokec4 (\cf7 \strokec7 edifData\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 0\cf4 \strokec4 ]) === \cf6 \strokec6 String\cf4 \strokec4 (\cf7 \strokec7 datos\cf4 \strokec4 .\cf7 \strokec7 idEdificio\cf4 \strokec4 )) \{ \cf7 \strokec7 parentFolderId\cf4 \strokec4  = \cf7 \strokec7 edifData\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 5\cf4 \strokec4 ]; \cf5 \strokec5 break\cf4 \strokec4 ; \} \cb1 \
\cb3   \}\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 assetFolderId\cf4 \strokec4  = \cf7 \strokec7 crearCarpeta\cf4 \strokec4 (\cf7 \strokec7 datos\cf4 \strokec4 .\cf7 \strokec7 nombre\cf4 \strokec4 , \cf7 \strokec7 parentFolderId\cf4 \strokec4 );\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 idActivo\cf4 \strokec4  = \cf6 \strokec6 Utilities\cf4 \strokec4 .\cf7 \strokec7 getUuid\cf4 \strokec4 ();\cb1 \
\cb3   \cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 sheetCat\cf4 \strokec4  = \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 getSheetByName\cf4 \strokec4 (\cf8 \strokec8 'CAT_INSTALACIONES'\cf4 \strokec4 );\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 catData\cf4 \strokec4  = \cf7 \strokec7 sheetCat\cf4 \strokec4 .\cf7 \strokec7 getDataRange\cf4 \strokec4 ().\cf7 \strokec7 getValues\cf4 \strokec4 ();\cb1 \
\cb3   \cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 diasRevision\cf4 \strokec4  = \cf9 \strokec9 365\cf4 \strokec4 ;\cb1 \
\cb3   \cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 nombreTipo\cf4 \strokec4  = \cf7 \strokec7 datos\cf4 \strokec4 .\cf7 \strokec7 tipo\cf4 \strokec4 ; \cb1 \
\
\cb3   \cf5 \strokec5 for\cf4 \strokec4 (\cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 i\cf4 \strokec4 =\cf9 \strokec9 1\cf4 \strokec4 ; \cf7 \strokec7 i\cf4 \strokec4 <\cf7 \strokec7 catData\cf4 \strokec4 .\cf7 \strokec7 length\cf4 \strokec4 ; \cf7 \strokec7 i\cf4 \strokec4 ++)\{\cb1 \
\cb3     \cf5 \strokec5 if\cf4 \strokec4 (\cf6 \strokec6 String\cf4 \strokec4 (\cf7 \strokec7 catData\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 0\cf4 \strokec4 ]) === \cf6 \strokec6 String\cf4 \strokec4 (\cf7 \strokec7 datos\cf4 \strokec4 .\cf7 \strokec7 tipo\cf4 \strokec4 )) \{\cb1 \
\cb3        \cf7 \strokec7 nombreTipo\cf4 \strokec4  = \cf7 \strokec7 catData\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 1\cf4 \strokec4 ];\cb1 \
\cb3        \cf7 \strokec7 diasRevision\cf4 \strokec4  = \cf7 \strokec7 catData\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 3\cf4 \strokec4 ];\cb1 \
\cb3        \cf5 \strokec5 break\cf4 \strokec4 ;\cb1 \
\cb3     \}\cb1 \
\cb3   \}\cb1 \
\
\cb3   \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 getSheetByName\cf4 \strokec4 (\cf8 \strokec8 'ACTIVOS'\cf4 \strokec4 ).\cf7 \strokec7 appendRow\cf4 \strokec4 ([\cf7 \strokec7 idActivo\cf4 \strokec4 , \cf7 \strokec7 datos\cf4 \strokec4 .\cf7 \strokec7 idEdificio\cf4 \strokec4 , \cf7 \strokec7 nombreTipo\cf4 \strokec4 , \cf7 \strokec7 datos\cf4 \strokec4 .\cf7 \strokec7 nombre\cf4 \strokec4 , \cf7 \strokec7 datos\cf4 \strokec4 .\cf7 \strokec7 marca\cf4 \strokec4 , \cf5 \strokec5 new\cf4 \strokec4  \cf6 \strokec6 Date\cf4 \strokec4 (), \cf7 \strokec7 assetFolderId\cf4 \strokec4 ]);\cb1 \
\cb3   \cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 sheetPlan\cf4 \strokec4  = \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 getSheetByName\cf4 \strokec4 (\cf8 \strokec8 'PLAN_MANTENIMIENTO'\cf4 \strokec4 );\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 fechaAlta\cf4 \strokec4  = \cf5 \strokec5 new\cf4 \strokec4  \cf6 \strokec6 Date\cf4 \strokec4 ();\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 fechaProx\cf4 \strokec4  = \cf5 \strokec5 new\cf4 \strokec4  \cf6 \strokec6 Date\cf4 \strokec4 (\cf7 \strokec7 fechaAlta\cf4 \strokec4 .\cf7 \strokec7 getTime\cf4 \strokec4 () + (\cf7 \strokec7 diasRevision\cf4 \strokec4  * \cf9 \strokec9 24\cf4 \strokec4  * \cf9 \strokec9 60\cf4 \strokec4  * \cf9 \strokec9 60\cf4 \strokec4  * \cf9 \strokec9 1000\cf4 \strokec4 ));\cb1 \
\cb3   \cb1 \
\cb3   \cf7 \strokec7 sheetPlan\cf4 \strokec4 .\cf7 \strokec7 appendRow\cf4 \strokec4 ([\cf6 \strokec6 Utilities\cf4 \strokec4 .\cf7 \strokec7 getUuid\cf4 \strokec4 (), \cf7 \strokec7 idActivo\cf4 \strokec4 , \cf8 \strokec8 `Mantenimiento Legal (\cf4 \strokec4 $\{\cf7 \strokec7 nombreTipo\cf4 \strokec4 \}\cf8 \strokec8 )`\cf4 \strokec4 , \cf8 \strokec8 ""\cf4 \strokec4 , \cf7 \strokec7 fechaProx\cf4 \strokec4 , \cf7 \strokec7 diasRevision\cf4 \strokec4 , \cf8 \strokec8 "ACTIVO"\cf4 \strokec4 ]);\cb1 \
\cb3   \cf5 \strokec5 return\cf4 \strokec4  \{ \cf7 \strokec7 success\cf4 \strokec4 : \cf5 \strokec5 true\cf4 \strokec4  \};\cb1 \
\cb3 \}\cb1 \
\
\pard\pardeftab720\partightenfactor0
\cf2 \cb3 \strokec2 // ==========================================\cf4 \cb1 \strokec4 \
\cf2 \cb3 \strokec2 // 6. GESTI\'d3N DOCUMENTAL\cf4 \cb1 \strokec4 \
\cf2 \cb3 \strokec2 // ==========================================\cf4 \cb1 \strokec4 \
\pard\pardeftab720\partightenfactor0
\cf5 \cb3 \strokec5 function\cf4 \strokec4  \cf7 \strokec7 subirArchivo\cf4 \strokec4 (\cf7 \strokec7 base64\cf4 \strokec4 , \cf7 \strokec7 nombre\cf4 \strokec4 , \cf7 \strokec7 mime\cf4 \strokec4 , \cf7 \strokec7 idEntidad\cf4 \strokec4 , \cf7 \strokec7 tipoEntidad\cf4 \strokec4 ) \{\cb1 \
\pard\pardeftab720\partightenfactor0
\cf4 \cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 lock\cf4 \strokec4  = \cf6 \strokec6 LockService\cf4 \strokec4 .\cf7 \strokec7 getScriptLock\cf4 \strokec4 ();\cb1 \
\cb3   \cf5 \strokec5 try\cf4 \strokec4  \{\cb1 \
\cb3     \cf7 \strokec7 lock\cf4 \strokec4 .\cf7 \strokec7 waitLock\cf4 \strokec4 (\cf9 \strokec9 10000\cf4 \strokec4 );\cb1 \
\cb3     \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 ss\cf4 \strokec4  = \cf6 \strokec6 SpreadsheetApp\cf4 \strokec4 .\cf7 \strokec7 openById\cf4 \strokec4 (\cf6 \strokec6 PROPS\cf4 \strokec4 .\cf7 \strokec7 getProperty\cf4 \strokec4 (\cf8 \strokec8 'DB_SS_ID'\cf4 \strokec4 ));\cb1 \
\cb3     \cb1 \
\cb3     \cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 targetFolderId\cf4 \strokec4  = \cf5 \strokec5 null\cf4 \strokec4 ;\cb1 \
\cb3     \cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 sheetName\cf4 \strokec4  = (\cf7 \strokec7 tipoEntidad\cf4 \strokec4  === \cf8 \strokec8 'EDIFICIO'\cf4 \strokec4 ) ? \cf8 \strokec8 'EDIFICIOS'\cf4 \strokec4  : \cf8 \strokec8 'ACTIVOS'\cf4 \strokec4 ;\cb1 \
\cb3     \cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 colIndex\cf4 \strokec4  = (\cf7 \strokec7 tipoEntidad\cf4 \strokec4  === \cf8 \strokec8 'EDIFICIO'\cf4 \strokec4 ) ? \cf9 \strokec9 4\cf4 \strokec4  : \cf9 \strokec9 6\cf4 \strokec4 ;\cb1 \
\cb3     \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 entityData\cf4 \strokec4  = \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 getSheetByName\cf4 \strokec4 (\cf7 \strokec7 sheetName\cf4 \strokec4 ).\cf7 \strokec7 getDataRange\cf4 \strokec4 ().\cf7 \strokec7 getValues\cf4 \strokec4 ();\cb1 \
\cb3     \cf5 \strokec5 for\cf4 \strokec4 (\cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 i\cf4 \strokec4 =\cf9 \strokec9 1\cf4 \strokec4 ; \cf7 \strokec7 i\cf4 \strokec4 <\cf7 \strokec7 entityData\cf4 \strokec4 .\cf7 \strokec7 length\cf4 \strokec4 ; \cf7 \strokec7 i\cf4 \strokec4 ++)\{\cb1 \
\cb3       \cf5 \strokec5 if\cf4 \strokec4 (\cf6 \strokec6 String\cf4 \strokec4 (\cf7 \strokec7 entityData\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 0\cf4 \strokec4 ]) === \cf6 \strokec6 String\cf4 \strokec4 (\cf7 \strokec7 idEntidad\cf4 \strokec4 )) \{ \cf7 \strokec7 targetFolderId\cf4 \strokec4  = \cf7 \strokec7 entityData\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf7 \strokec7 colIndex\cf4 \strokec4 ]; \cf5 \strokec5 break\cf4 \strokec4 ; \}\cb1 \
\cb3     \}\cb1 \
\cb3     \cb1 \
\cb3     \cf5 \strokec5 if\cf4 \strokec4 (!\cf7 \strokec7 targetFolderId\cf4 \strokec4 ) \cf5 \strokec5 throw\cf4 \strokec4  \cf5 \strokec5 new\cf4 \strokec4  \cf6 \strokec6 Error\cf4 \strokec4 (\cf8 \strokec8 "Carpeta no encontrada"\cf4 \strokec4 );\cb1 \
\cb3     \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 sheetDocs\cf4 \strokec4  = \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 getSheetByName\cf4 \strokec4 (\cf8 \strokec8 'DOCS_HISTORICO'\cf4 \strokec4 );\cb1 \
\cb3     \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 docsData\cf4 \strokec4  = \cf7 \strokec7 sheetDocs\cf4 \strokec4 .\cf7 \strokec7 getDataRange\cf4 \strokec4 ().\cf7 \strokec7 getValues\cf4 \strokec4 ();\cb1 \
\cb3     \cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 maxVer\cf4 \strokec4  = \cf9 \strokec9 0\cf4 \strokec4 ;\cb1 \
\cb3     \cf5 \strokec5 for\cf4 \strokec4 (\cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 i\cf4 \strokec4 =\cf9 \strokec9 1\cf4 \strokec4 ; \cf7 \strokec7 i\cf4 \strokec4 <\cf7 \strokec7 docsData\cf4 \strokec4 .\cf7 \strokec7 length\cf4 \strokec4 ; \cf7 \strokec7 i\cf4 \strokec4 ++)\{\cb1 \
\cb3       \cf5 \strokec5 if\cf4 \strokec4 (\cf6 \strokec6 String\cf4 \strokec4 (\cf7 \strokec7 docsData\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 2\cf4 \strokec4 ]) === \cf6 \strokec6 String\cf4 \strokec4 (\cf7 \strokec7 idEntidad\cf4 \strokec4 ) && \cf7 \strokec7 docsData\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 3\cf4 \strokec4 ] === \cf7 \strokec7 nombre\cf4 \strokec4 ) \{\cb1 \
\cb3         \cf5 \strokec5 if\cf4 \strokec4 (\cf7 \strokec7 docsData\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 5\cf4 \strokec4 ] > \cf7 \strokec7 maxVer\cf4 \strokec4 ) \cf7 \strokec7 maxVer\cf4 \strokec4  = \cf7 \strokec7 docsData\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 5\cf4 \strokec4 ];\cb1 \
\cb3       \}\cb1 \
\cb3     \}\cb1 \
\cb3     \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 newVer\cf4 \strokec4  = \cf7 \strokec7 maxVer\cf4 \strokec4  + \cf9 \strokec9 1\cf4 \strokec4 ;\cb1 \
\cb3     \cb1 \
\cb3     \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 blob\cf4 \strokec4  = \cf6 \strokec6 Utilities\cf4 \strokec4 .\cf7 \strokec7 newBlob\cf4 \strokec4 (\cf6 \strokec6 Utilities\cf4 \strokec4 .\cf7 \strokec7 base64Decode\cf4 \strokec4 (\cf7 \strokec7 base64\cf4 \strokec4 ), \cf7 \strokec7 mime\cf4 \strokec4 , \cf8 \strokec8 `[v\cf4 \strokec4 $\{\cf7 \strokec7 newVer\cf4 \strokec4 \}\cf8 \strokec8 ] \cf4 \strokec4 $\{\cf7 \strokec7 nombre\cf4 \strokec4 \}\cf8 \strokec8 `\cf4 \strokec4 );\cb1 \
\cb3     \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 file\cf4 \strokec4  = \cf6 \strokec6 DriveApp\cf4 \strokec4 .\cf7 \strokec7 getFolderById\cf4 \strokec4 (\cf7 \strokec7 targetFolderId\cf4 \strokec4 ).\cf7 \strokec7 createFile\cf4 \strokec4 (\cf7 \strokec7 blob\cf4 \strokec4 );\cb1 \
\cb3     \cb1 \
\cb3     \cf7 \strokec7 sheetDocs\cf4 \strokec4 .\cf7 \strokec7 appendRow\cf4 \strokec4 ([\cf6 \strokec6 Utilities\cf4 \strokec4 .\cf7 \strokec7 getUuid\cf4 \strokec4 (), \cf7 \strokec7 tipoEntidad\cf4 \strokec4 , \cf7 \strokec7 idEntidad\cf4 \strokec4 , \cf7 \strokec7 nombre\cf4 \strokec4 , \cf7 \strokec7 file\cf4 \strokec4 .\cf7 \strokec7 getUrl\cf4 \strokec4 (), \cf7 \strokec7 newVer\cf4 \strokec4 , \cf5 \strokec5 new\cf4 \strokec4  \cf6 \strokec6 Date\cf4 \strokec4 (), \cf6 \strokec6 Session\cf4 \strokec4 .\cf7 \strokec7 getActiveUser\cf4 \strokec4 ().\cf7 \strokec7 getEmail\cf4 \strokec4 (), \cf7 \strokec7 file\cf4 \strokec4 .\cf7 \strokec7 getId\cf4 \strokec4 ()]);\cb1 \
\cb3     \cb1 \
\cb3     \cf5 \strokec5 return\cf4 \strokec4  \{ \cf7 \strokec7 success\cf4 \strokec4 : \cf5 \strokec5 true\cf4 \strokec4 , \cf7 \strokec7 version\cf4 \strokec4 : \cf7 \strokec7 newVer\cf4 \strokec4  \};\cb1 \
\cb3   \} \cf5 \strokec5 catch\cf4 \strokec4 (\cf7 \strokec7 e\cf4 \strokec4 ) \{ \cf5 \strokec5 return\cf4 \strokec4  \{ \cf7 \strokec7 success\cf4 \strokec4 : \cf5 \strokec5 false\cf4 \strokec4 , \cf7 \strokec7 error\cf4 \strokec4 : \cf7 \strokec7 e\cf4 \strokec4 .\cf7 \strokec7 toString\cf4 \strokec4 () \}; \} \cf5 \strokec5 finally\cf4 \strokec4  \{ \cf7 \strokec7 lock\cf4 \strokec4 .\cf7 \strokec7 releaseLock\cf4 \strokec4 (); \}\cb1 \
\cb3 \}\cb1 \
\
\pard\pardeftab720\partightenfactor0
\cf5 \cb3 \strokec5 function\cf4 \strokec4  \cf7 \strokec7 obtenerDocs\cf4 \strokec4 (\cf7 \strokec7 idEntidad\cf4 \strokec4 ) \{\cb1 \
\pard\pardeftab720\partightenfactor0
\cf4 \cb3   \cf5 \strokec5 try\cf4 \strokec4  \{\cb1 \
\cb3     \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 ss\cf4 \strokec4  = \cf6 \strokec6 SpreadsheetApp\cf4 \strokec4 .\cf7 \strokec7 openById\cf4 \strokec4 (\cf6 \strokec6 PROPS\cf4 \strokec4 .\cf7 \strokec7 getProperty\cf4 \strokec4 (\cf8 \strokec8 'DB_SS_ID'\cf4 \strokec4 ));\cb1 \
\cb3     \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 sheet\cf4 \strokec4  = \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 getSheetByName\cf4 \strokec4 (\cf8 \strokec8 'DOCS_HISTORICO'\cf4 \strokec4 );\cb1 \
\cb3     \cf5 \strokec5 if\cf4 \strokec4  (\cf7 \strokec7 sheet\cf4 \strokec4 .\cf7 \strokec7 getLastRow\cf4 \strokec4 () < \cf9 \strokec9 2\cf4 \strokec4 ) \cf5 \strokec5 return\cf4 \strokec4  [];\cb1 \
\cb3     \cb1 \
\cb3     \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 data\cf4 \strokec4  = \cf7 \strokec7 sheet\cf4 \strokec4 .\cf7 \strokec7 getDataRange\cf4 \strokec4 ().\cf7 \strokec7 getValues\cf4 \strokec4 ();\cb1 \
\cb3     \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 res\cf4 \strokec4  = [];\cb1 \
\cb3     \cf5 \strokec5 for\cf4 \strokec4 (\cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 i\cf4 \strokec4  = \cf7 \strokec7 data\cf4 \strokec4 .\cf7 \strokec7 length\cf4 \strokec4  - \cf9 \strokec9 1\cf4 \strokec4 ; \cf7 \strokec7 i\cf4 \strokec4  >= \cf9 \strokec9 1\cf4 \strokec4 ; \cf7 \strokec7 i\cf4 \strokec4 --)\{\cb1 \
\cb3       \cf5 \strokec5 if\cf4 \strokec4 (\cf6 \strokec6 String\cf4 \strokec4 (\cf7 \strokec7 data\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 2\cf4 \strokec4 ]) === \cf6 \strokec6 String\cf4 \strokec4 (\cf7 \strokec7 idEntidad\cf4 \strokec4 )) \{\cb1 \
\cb3         \cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 fechaFormateada\cf4 \strokec4  = \cf8 \strokec8 ""\cf4 \strokec4 ;\cb1 \
\cb3         \cf5 \strokec5 try\cf4 \strokec4  \{\cb1 \
\cb3           \cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 fechaRaw\cf4 \strokec4  = \cf7 \strokec7 data\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 6\cf4 \strokec4 ];\cb1 \
\cb3           \cf7 \strokec7 fechaFormateada\cf4 \strokec4  = (\cf7 \strokec7 fechaRaw\cf4 \strokec4  \cf5 \strokec5 instanceof\cf4 \strokec4  \cf6 \strokec6 Date\cf4 \strokec4 ) ? \cf6 \strokec6 Utilities\cf4 \strokec4 .\cf7 \strokec7 formatDate\cf4 \strokec4 (\cf7 \strokec7 fechaRaw\cf4 \strokec4 , \cf6 \strokec6 Session\cf4 \strokec4 .\cf7 \strokec7 getScriptTimeZone\cf4 \strokec4 (), \cf8 \strokec8 "dd/MM/yyyy"\cf4 \strokec4 ) : \cf6 \strokec6 String\cf4 \strokec4 (\cf7 \strokec7 fechaRaw\cf4 \strokec4 );\cb1 \
\cb3         \} \cf5 \strokec5 catch\cf4 \strokec4 (\cf7 \strokec7 e\cf4 \strokec4 ) \{ \cf7 \strokec7 fechaFormateada\cf4 \strokec4  = \cf8 \strokec8 "--"\cf4 \strokec4 ; \}\cb1 \
\cb3         \cf7 \strokec7 res\cf4 \strokec4 .\cf7 \strokec7 push\cf4 \strokec4 (\{ \cf7 \strokec7 nombre\cf4 \strokec4 : \cf7 \strokec7 data\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 3\cf4 \strokec4 ], \cf7 \strokec7 url\cf4 \strokec4 : \cf7 \strokec7 data\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 4\cf4 \strokec4 ], \cf7 \strokec7 version\cf4 \strokec4 : \cf7 \strokec7 data\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 5\cf4 \strokec4 ], \cf7 \strokec7 fecha\cf4 \strokec4 : \cf7 \strokec7 fechaFormateada\cf4 \strokec4  \});\cb1 \
\cb3       \}\cb1 \
\cb3     \}\cb1 \
\cb3     \cf5 \strokec5 return\cf4 \strokec4  \cf7 \strokec7 res\cf4 \strokec4 ;\cb1 \
\cb3   \} \cf5 \strokec5 catch\cf4 \strokec4  (\cf7 \strokec7 e\cf4 \strokec4 ) \{ \cf5 \strokec5 throw\cf4 \strokec4  \cf5 \strokec5 new\cf4 \strokec4  \cf6 \strokec6 Error\cf4 \strokec4 (\cf7 \strokec7 e\cf4 \strokec4 .\cf7 \strokec7 toString\cf4 \strokec4 ()); \}\cb1 \
\cb3 \}\cb1 \
\
\pard\pardeftab720\partightenfactor0
\cf2 \cb3 \strokec2 // ==========================================\cf4 \cb1 \strokec4 \
\cf2 \cb3 \strokec2 // 7. GESTI\'d3N MANTENIMIENTO\cf4 \cb1 \strokec4 \
\cf2 \cb3 \strokec2 // ==========================================\cf4 \cb1 \strokec4 \
\pard\pardeftab720\partightenfactor0
\cf5 \cb3 \strokec5 function\cf4 \strokec4  \cf7 \strokec7 obtenerPlanMantenimiento\cf4 \strokec4 (\cf7 \strokec7 idActivo\cf4 \strokec4 ) \{\cb1 \
\pard\pardeftab720\partightenfactor0
\cf4 \cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 ss\cf4 \strokec4  = \cf6 \strokec6 SpreadsheetApp\cf4 \strokec4 .\cf7 \strokec7 openById\cf4 \strokec4 (\cf6 \strokec6 PROPS\cf4 \strokec4 .\cf7 \strokec7 getProperty\cf4 \strokec4 (\cf8 \strokec8 'DB_SS_ID'\cf4 \strokec4 ));\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 sheet\cf4 \strokec4  = \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 getSheetByName\cf4 \strokec4 (\cf8 \strokec8 'PLAN_MANTENIMIENTO'\cf4 \strokec4 );\cb1 \
\cb3   \cf5 \strokec5 if\cf4 \strokec4 (!\cf7 \strokec7 sheet\cf4 \strokec4  || \cf7 \strokec7 sheet\cf4 \strokec4 .\cf7 \strokec7 getLastRow\cf4 \strokec4 () < \cf9 \strokec9 2\cf4 \strokec4 ) \cf5 \strokec5 return\cf4 \strokec4  [];\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 data\cf4 \strokec4  = \cf7 \strokec7 sheet\cf4 \strokec4 .\cf7 \strokec7 getDataRange\cf4 \strokec4 ().\cf7 \strokec7 getValues\cf4 \strokec4 ();\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 planes\cf4 \strokec4  = [];\cb1 \
\cb3   \cf5 \strokec5 for\cf4 \strokec4 (\cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 i\cf4 \strokec4 =\cf9 \strokec9 1\cf4 \strokec4 ; \cf7 \strokec7 i\cf4 \strokec4 <\cf7 \strokec7 data\cf4 \strokec4 .\cf7 \strokec7 length\cf4 \strokec4 ; \cf7 \strokec7 i\cf4 \strokec4 ++) \{\cb1 \
\cb3     \cf5 \strokec5 if\cf4 \strokec4 (\cf6 \strokec6 String\cf4 \strokec4 (\cf7 \strokec7 data\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 1\cf4 \strokec4 ]) === \cf6 \strokec6 String\cf4 \strokec4 (\cf7 \strokec7 idActivo\cf4 \strokec4 )) \{\cb1 \
\cb3       \cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 fProx\cf4 \strokec4  = \cf7 \strokec7 data\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 4\cf4 \strokec4 ];\cb1 \
\cb3       \cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 fechaStr\cf4 \strokec4  = (\cf7 \strokec7 fProx\cf4 \strokec4  \cf5 \strokec5 instanceof\cf4 \strokec4  \cf6 \strokec6 Date\cf4 \strokec4 ) ? \cf6 \strokec6 Utilities\cf4 \strokec4 .\cf7 \strokec7 formatDate\cf4 \strokec4 (\cf7 \strokec7 fProx\cf4 \strokec4 , \cf6 \strokec6 Session\cf4 \strokec4 .\cf7 \strokec7 getScriptTimeZone\cf4 \strokec4 (), \cf8 \strokec8 "yyyy-MM-dd"\cf4 \strokec4 ) : \cf8 \strokec8 ""\cf4 \strokec4 ;\cb1 \
\cb3       \cf7 \strokec7 planes\cf4 \strokec4 .\cf7 \strokec7 push\cf4 \strokec4 (\{ \cf7 \strokec7 id\cf4 \strokec4 : \cf7 \strokec7 data\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 0\cf4 \strokec4 ], \cf7 \strokec7 tipo\cf4 \strokec4 : \cf7 \strokec7 data\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 2\cf4 \strokec4 ], \cf7 \strokec7 fechaProxima\cf4 \strokec4 : \cf7 \strokec7 fechaStr\cf4 \strokec4  \});\cb1 \
\cb3     \}\cb1 \
\cb3   \}\cb1 \
\cb3   \cf5 \strokec5 return\cf4 \strokec4  \cf7 \strokec7 planes\cf4 \strokec4 ;\cb1 \
\cb3 \}\cb1 \
\
\pard\pardeftab720\partightenfactor0
\cf5 \cb3 \strokec5 function\cf4 \strokec4  \cf7 \strokec7 crearRevision\cf4 \strokec4 (\cf7 \strokec7 datos\cf4 \strokec4 ) \{ \cb1 \
\pard\pardeftab720\partightenfactor0
\cf4 \cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 ss\cf4 \strokec4  = \cf6 \strokec6 SpreadsheetApp\cf4 \strokec4 .\cf7 \strokec7 openById\cf4 \strokec4 (\cf6 \strokec6 PROPS\cf4 \strokec4 .\cf7 \strokec7 getProperty\cf4 \strokec4 (\cf8 \strokec8 'DB_SS_ID'\cf4 \strokec4 ));\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 sheet\cf4 \strokec4  = \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 getSheetByName\cf4 \strokec4 (\cf8 \strokec8 'PLAN_MANTENIMIENTO'\cf4 \strokec4 );\cb1 \
\cb3   \cf7 \strokec7 sheet\cf4 \strokec4 .\cf7 \strokec7 appendRow\cf4 \strokec4 ([\cf6 \strokec6 Utilities\cf4 \strokec4 .\cf7 \strokec7 getUuid\cf4 \strokec4 (), \cf7 \strokec7 datos\cf4 \strokec4 .\cf7 \strokec7 idActivo\cf4 \strokec4 , \cf7 \strokec7 datos\cf4 \strokec4 .\cf7 \strokec7 tipo\cf4 \strokec4 , \cf8 \strokec8 ""\cf4 \strokec4 , \cf5 \strokec5 new\cf4 \strokec4  \cf6 \strokec6 Date\cf4 \strokec4 (\cf7 \strokec7 datos\cf4 \strokec4 .\cf7 \strokec7 fechaProx\cf4 \strokec4 ), \cf9 \strokec9 365\cf4 \strokec4 , \cf8 \strokec8 "ACTIVO"\cf4 \strokec4 ]);\cb1 \
\cb3   \cf5 \strokec5 return\cf4 \strokec4  \{ \cf7 \strokec7 success\cf4 \strokec4 : \cf5 \strokec5 true\cf4 \strokec4  \};\cb1 \
\cb3 \}\cb1 \
\
\pard\pardeftab720\partightenfactor0
\cf2 \cb3 \strokec2 // ==========================================\cf4 \cb1 \strokec4 \
\cf2 \cb3 \strokec2 // 8. GESTI\'d3N CONTRATOS\cf4 \cb1 \strokec4 \
\cf2 \cb3 \strokec2 // ==========================================\cf4 \cb1 \strokec4 \
\pard\pardeftab720\partightenfactor0
\cf5 \cb3 \strokec5 function\cf4 \strokec4  \cf7 \strokec7 obtenerContratos\cf4 \strokec4 (\cf7 \strokec7 idEntidad\cf4 \strokec4 ) \{\cb1 \
\pard\pardeftab720\partightenfactor0
\cf4 \cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 ss\cf4 \strokec4  = \cf6 \strokec6 SpreadsheetApp\cf4 \strokec4 .\cf7 \strokec7 openById\cf4 \strokec4 (\cf6 \strokec6 PROPS\cf4 \strokec4 .\cf7 \strokec7 getProperty\cf4 \strokec4 (\cf8 \strokec8 'DB_SS_ID'\cf4 \strokec4 ));\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 sheet\cf4 \strokec4  = \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 getSheetByName\cf4 \strokec4 (\cf8 \strokec8 'CONTRATOS'\cf4 \strokec4 );\cb1 \
\cb3   \cf5 \strokec5 if\cf4 \strokec4 (!\cf7 \strokec7 sheet\cf4 \strokec4  || \cf7 \strokec7 sheet\cf4 \strokec4 .\cf7 \strokec7 getLastRow\cf4 \strokec4 () < \cf9 \strokec9 2\cf4 \strokec4 ) \cf5 \strokec5 return\cf4 \strokec4  [];\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 data\cf4 \strokec4  = \cf7 \strokec7 sheet\cf4 \strokec4 .\cf7 \strokec7 getDataRange\cf4 \strokec4 ().\cf7 \strokec7 getValues\cf4 \strokec4 ();\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 contratos\cf4 \strokec4  = [];\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 hoy\cf4 \strokec4  = \cf5 \strokec5 new\cf4 \strokec4  \cf6 \strokec6 Date\cf4 \strokec4 ();\cb1 \
\cb3   \cf5 \strokec5 for\cf4 \strokec4 (\cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 i\cf4 \strokec4 =\cf9 \strokec9 1\cf4 \strokec4 ; \cf7 \strokec7 i\cf4 \strokec4 <\cf7 \strokec7 data\cf4 \strokec4 .\cf7 \strokec7 length\cf4 \strokec4 ; \cf7 \strokec7 i\cf4 \strokec4 ++) \{\cb1 \
\cb3     \cf5 \strokec5 if\cf4 \strokec4 (\cf6 \strokec6 String\cf4 \strokec4 (\cf7 \strokec7 data\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 2\cf4 \strokec4 ]) === \cf6 \strokec6 String\cf4 \strokec4 (\cf7 \strokec7 idEntidad\cf4 \strokec4 )) \{\cb1 \
\cb3       \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 fFin\cf4 \strokec4  = \cf7 \strokec7 data\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 6\cf4 \strokec4 ] \cf5 \strokec5 instanceof\cf4 \strokec4  \cf6 \strokec6 Date\cf4 \strokec4  ? \cf7 \strokec7 data\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 6\cf4 \strokec4 ] : \cf5 \strokec5 new\cf4 \strokec4  \cf6 \strokec6 Date\cf4 \strokec4 (\cf7 \strokec7 data\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 6\cf4 \strokec4 ]);\cb1 \
\cb3       \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 fIni\cf4 \strokec4  = \cf7 \strokec7 data\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 5\cf4 \strokec4 ] \cf5 \strokec5 instanceof\cf4 \strokec4  \cf6 \strokec6 Date\cf4 \strokec4  ? \cf7 \strokec7 data\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 5\cf4 \strokec4 ] : \cf5 \strokec5 new\cf4 \strokec4  \cf6 \strokec6 Date\cf4 \strokec4 (\cf7 \strokec7 data\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 5\cf4 \strokec4 ]);\cb1 \
\cb3       \cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 estado\cf4 \strokec4  = \cf8 \strokec8 'VIGENTE'\cf4 \strokec4 ; \cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 color\cf4 \strokec4  = \cf8 \strokec8 'verde'\cf4 \strokec4 ;\cb1 \
\cb3       \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 diffDays\cf4 \strokec4  = \cf6 \strokec6 Math\cf4 \strokec4 .\cf7 \strokec7 ceil\cf4 \strokec4 ((\cf7 \strokec7 fFin\cf4 \strokec4  - \cf7 \strokec7 hoy\cf4 \strokec4 ) / (\cf9 \strokec9 1000\cf4 \strokec4  * \cf9 \strokec9 60\cf4 \strokec4  * \cf9 \strokec9 60\cf4 \strokec4  * \cf9 \strokec9 24\cf4 \strokec4 ));\cb1 \
\cb3       \cf5 \strokec5 if\cf4 \strokec4  (\cf7 \strokec7 diffDays\cf4 \strokec4  < \cf9 \strokec9 0\cf4 \strokec4 ) \{ \cf7 \strokec7 estado\cf4 \strokec4  = \cf8 \strokec8 'CADUCADO'\cf4 \strokec4 ; \cf7 \strokec7 color\cf4 \strokec4  = \cf8 \strokec8 'rojo'\cf4 \strokec4 ; \}\cb1 \
\cb3       \cf5 \strokec5 else\cf4 \strokec4  \cf5 \strokec5 if\cf4 \strokec4  (\cf7 \strokec7 diffDays\cf4 \strokec4  <= \cf9 \strokec9 30\cf4 \strokec4 ) \{ \cf7 \strokec7 estado\cf4 \strokec4  = \cf8 \strokec8 'PR\'d3XIMO'\cf4 \strokec4 ; \cf7 \strokec7 color\cf4 \strokec4  = \cf8 \strokec8 'amarillo'\cf4 \strokec4 ; \}\cb1 \
\cb3       \cf7 \strokec7 contratos\cf4 \strokec4 .\cf7 \strokec7 push\cf4 \strokec4 (\{\cb1 \
\cb3         \cf7 \strokec7 id\cf4 \strokec4 : \cf7 \strokec7 data\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 0\cf4 \strokec4 ], \cf7 \strokec7 proveedor\cf4 \strokec4 : \cf7 \strokec7 data\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 3\cf4 \strokec4 ], \cf7 \strokec7 ref\cf4 \strokec4 : \cf7 \strokec7 data\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 4\cf4 \strokec4 ],\cb1 \
\cb3         \cf7 \strokec7 inicio\cf4 \strokec4 : \cf6 \strokec6 Utilities\cf4 \strokec4 .\cf7 \strokec7 formatDate\cf4 \strokec4 (\cf7 \strokec7 fIni\cf4 \strokec4 , \cf6 \strokec6 Session\cf4 \strokec4 .\cf7 \strokec7 getScriptTimeZone\cf4 \strokec4 (), \cf8 \strokec8 "dd/MM/yyyy"\cf4 \strokec4 ),\cb1 \
\cb3         \cf7 \strokec7 fin\cf4 \strokec4 : \cf6 \strokec6 Utilities\cf4 \strokec4 .\cf7 \strokec7 formatDate\cf4 \strokec4 (\cf7 \strokec7 fFin\cf4 \strokec4 , \cf6 \strokec6 Session\cf4 \strokec4 .\cf7 \strokec7 getScriptTimeZone\cf4 \strokec4 (), \cf8 \strokec8 "dd/MM/yyyy"\cf4 \strokec4 ),\cb1 \
\cb3         \cf7 \strokec7 estado\cf4 \strokec4 : \cf7 \strokec7 estado\cf4 \strokec4 , \cf7 \strokec7 color\cf4 \strokec4 : \cf7 \strokec7 color\cf4 \cb1 \strokec4 \
\cb3       \});\cb1 \
\cb3     \}\cb1 \
\cb3   \}\cb1 \
\cb3   \cf5 \strokec5 return\cf4 \strokec4  \cf7 \strokec7 contratos\cf4 \strokec4 ;\cb1 \
\cb3 \}\cb1 \
\
\pard\pardeftab720\partightenfactor0
\cf5 \cb3 \strokec5 function\cf4 \strokec4  \cf7 \strokec7 crearContrato\cf4 \strokec4 (\cf7 \strokec7 datos\cf4 \strokec4 ) \{\cb1 \
\pard\pardeftab720\partightenfactor0
\cf4 \cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 ss\cf4 \strokec4  = \cf6 \strokec6 SpreadsheetApp\cf4 \strokec4 .\cf7 \strokec7 openById\cf4 \strokec4 (\cf6 \strokec6 PROPS\cf4 \strokec4 .\cf7 \strokec7 getProperty\cf4 \strokec4 (\cf8 \strokec8 'DB_SS_ID'\cf4 \strokec4 ));\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 sheet\cf4 \strokec4  = \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 getSheetByName\cf4 \strokec4 (\cf8 \strokec8 'CONTRATOS'\cf4 \strokec4 );\cb1 \
\cb3   \cf7 \strokec7 sheet\cf4 \strokec4 .\cf7 \strokec7 appendRow\cf4 \strokec4 ([\cf6 \strokec6 Utilities\cf4 \strokec4 .\cf7 \strokec7 getUuid\cf4 \strokec4 (), \cf7 \strokec7 datos\cf4 \strokec4 .\cf7 \strokec7 tipoEntidad\cf4 \strokec4 , \cf7 \strokec7 datos\cf4 \strokec4 .\cf7 \strokec7 idEntidad\cf4 \strokec4 , \cf7 \strokec7 datos\cf4 \strokec4 .\cf7 \strokec7 proveedor\cf4 \strokec4 , \cf7 \strokec7 datos\cf4 \strokec4 .\cf7 \strokec7 ref\cf4 \strokec4 , \cf5 \strokec5 new\cf4 \strokec4  \cf6 \strokec6 Date\cf4 \strokec4 (\cf7 \strokec7 datos\cf4 \strokec4 .\cf7 \strokec7 fechaIni\cf4 \strokec4 ), \cf5 \strokec5 new\cf4 \strokec4  \cf6 \strokec6 Date\cf4 \strokec4 (\cf7 \strokec7 datos\cf4 \strokec4 .\cf7 \strokec7 fechaFin\cf4 \strokec4 )]);\cb1 \
\cb3   \cf5 \strokec5 return\cf4 \strokec4  \{ \cf7 \strokec7 success\cf4 \strokec4 : \cf5 \strokec5 true\cf4 \strokec4  \};\cb1 \
\cb3 \}\cb1 \
\
\pard\pardeftab720\partightenfactor0
\cf2 \cb3 \strokec2 // ==========================================\cf4 \cb1 \strokec4 \
\cf2 \cb3 \strokec2 // 9. AUTOMATIZACI\'d3N DE AVISOS Y DASHBOARD\cf4 \cb1 \strokec4 \
\cf2 \cb3 \strokec2 // ==========================================\cf4 \cb1 \strokec4 \
\pard\pardeftab720\partightenfactor0
\cf5 \cb3 \strokec5 function\cf4 \strokec4  \cf7 \strokec7 getDatosDashboard\cf4 \strokec4 () \{\cb1 \
\pard\pardeftab720\partightenfactor0
\cf4 \cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 ss\cf4 \strokec4  = \cf6 \strokec6 SpreadsheetApp\cf4 \strokec4 .\cf7 \strokec7 openById\cf4 \strokec4 (\cf6 \strokec6 PROPS\cf4 \strokec4 .\cf7 \strokec7 getProperty\cf4 \strokec4 (\cf8 \strokec8 'DB_SS_ID'\cf4 \strokec4 ));\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 hoy\cf4 \strokec4  = \cf5 \strokec5 new\cf4 \strokec4  \cf6 \strokec6 Date\cf4 \strokec4 ();\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 sheetAct\cf4 \strokec4  = \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 getSheetByName\cf4 \strokec4 (\cf8 \strokec8 'ACTIVOS'\cf4 \strokec4 );\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 totalActivos\cf4 \strokec4  = \cf7 \strokec7 sheetAct\cf4 \strokec4 .\cf7 \strokec7 getLastRow\cf4 \strokec4 () - \cf9 \strokec9 1\cf4 \strokec4 ; \cb1 \
\
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 sheetMant\cf4 \strokec4  = \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 getSheetByName\cf4 \strokec4 (\cf8 \strokec8 'PLAN_MANTENIMIENTO'\cf4 \strokec4 );\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 dataMant\cf4 \strokec4  = \cf7 \strokec7 sheetMant\cf4 \strokec4 .\cf7 \strokec7 getDataRange\cf4 \strokec4 ().\cf7 \strokec7 getValues\cf4 \strokec4 ();\cb1 \
\cb3   \cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 revPendientes\cf4 \strokec4  = \cf9 \strokec9 0\cf4 \strokec4 , \cf7 \strokec7 revOk\cf4 \strokec4  = \cf9 \strokec9 0\cf4 \strokec4 , \cf7 \strokec7 revVencidas\cf4 \strokec4  = \cf9 \strokec9 0\cf4 \strokec4 ;\cb1 \
\cb3   \cf5 \strokec5 for\cf4 \strokec4 (\cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 i\cf4 \strokec4 =\cf9 \strokec9 1\cf4 \strokec4 ; \cf7 \strokec7 i\cf4 \strokec4 <\cf7 \strokec7 dataMant\cf4 \strokec4 .\cf7 \strokec7 length\cf4 \strokec4 ; \cf7 \strokec7 i\cf4 \strokec4 ++) \{\cb1 \
\cb3     \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 fecha\cf4 \strokec4  = \cf7 \strokec7 dataMant\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 4\cf4 \strokec4 ] \cf5 \strokec5 instanceof\cf4 \strokec4  \cf6 \strokec6 Date\cf4 \strokec4  ? \cf7 \strokec7 dataMant\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 4\cf4 \strokec4 ] : \cf5 \strokec5 new\cf4 \strokec4  \cf6 \strokec6 Date\cf4 \strokec4 (\cf7 \strokec7 dataMant\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 4\cf4 \strokec4 ]);\cb1 \
\cb3     \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 diffDays\cf4 \strokec4  = \cf6 \strokec6 Math\cf4 \strokec4 .\cf7 \strokec7 ceil\cf4 \strokec4 ((\cf7 \strokec7 fecha\cf4 \strokec4  - \cf7 \strokec7 hoy\cf4 \strokec4 ) / (\cf9 \strokec9 1000\cf4 \strokec4  * \cf9 \strokec9 60\cf4 \strokec4  * \cf9 \strokec9 60\cf4 \strokec4  * \cf9 \strokec9 24\cf4 \strokec4 ));\cb1 \
\cb3     \cf5 \strokec5 if\cf4 \strokec4  (\cf7 \strokec7 diffDays\cf4 \strokec4  < \cf9 \strokec9 0\cf4 \strokec4 ) \cf7 \strokec7 revVencidas\cf4 \strokec4 ++; \cf5 \strokec5 else\cf4 \strokec4  \cf5 \strokec5 if\cf4 \strokec4  (\cf7 \strokec7 diffDays\cf4 \strokec4  <= \cf9 \strokec9 30\cf4 \strokec4 ) \cf7 \strokec7 revPendientes\cf4 \strokec4 ++; \cf5 \strokec5 else\cf4 \strokec4  \cf7 \strokec7 revOk\cf4 \strokec4 ++;\cb1 \
\cb3   \}\cb1 \
\
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 sheetCont\cf4 \strokec4  = \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 getSheetByName\cf4 \strokec4 (\cf8 \strokec8 'CONTRATOS'\cf4 \strokec4 );\cb1 \
\cb3   \cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 contCaducados\cf4 \strokec4  = \cf9 \strokec9 0\cf4 \strokec4 ;\cb1 \
\cb3   \cf5 \strokec5 if\cf4 \strokec4 (\cf7 \strokec7 sheetCont\cf4 \strokec4  && \cf7 \strokec7 sheetCont\cf4 \strokec4 .\cf7 \strokec7 getLastRow\cf4 \strokec4 () > \cf9 \strokec9 1\cf4 \strokec4 ) \{\cb1 \
\cb3     \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 dataCont\cf4 \strokec4  = \cf7 \strokec7 sheetCont\cf4 \strokec4 .\cf7 \strokec7 getDataRange\cf4 \strokec4 ().\cf7 \strokec7 getValues\cf4 \strokec4 ();\cb1 \
\cb3     \cf5 \strokec5 for\cf4 \strokec4 (\cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 i\cf4 \strokec4 =\cf9 \strokec9 1\cf4 \strokec4 ; \cf7 \strokec7 i\cf4 \strokec4 <\cf7 \strokec7 dataCont\cf4 \strokec4 .\cf7 \strokec7 length\cf4 \strokec4 ; \cf7 \strokec7 i\cf4 \strokec4 ++) \{\cb1 \
\cb3       \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 fecha\cf4 \strokec4  = \cf7 \strokec7 dataCont\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 6\cf4 \strokec4 ] \cf5 \strokec5 instanceof\cf4 \strokec4  \cf6 \strokec6 Date\cf4 \strokec4  ? \cf7 \strokec7 dataCont\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 6\cf4 \strokec4 ] : \cf5 \strokec5 new\cf4 \strokec4  \cf6 \strokec6 Date\cf4 \strokec4 (\cf7 \strokec7 dataCont\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 6\cf4 \strokec4 ]);\cb1 \
\cb3       \cf5 \strokec5 if\cf4 \strokec4  (\cf7 \strokec7 fecha\cf4 \strokec4  < \cf7 \strokec7 hoy\cf4 \strokec4 ) \cf7 \strokec7 contCaducados\cf4 \strokec4 ++;\cb1 \
\cb3     \}\cb1 \
\cb3   \}\cb1 \
\
\cb3   \cf5 \strokec5 return\cf4 \strokec4  \{ \cf7 \strokec7 activos\cf4 \strokec4 : \cf7 \strokec7 totalActivos\cf4 \strokec4  > \cf9 \strokec9 0\cf4 \strokec4  ? \cf7 \strokec7 totalActivos\cf4 \strokec4  : \cf9 \strokec9 0\cf4 \strokec4 , \cf7 \strokec7 pendientes\cf4 \strokec4 : \cf7 \strokec7 revPendientes\cf4 \strokec4 , \cf7 \strokec7 vencidas\cf4 \strokec4 : \cf7 \strokec7 revVencidas\cf4 \strokec4 , \cf7 \strokec7 ok\cf4 \strokec4 : \cf7 \strokec7 revOk\cf4 \strokec4 , \cf7 \strokec7 contratosCaducados\cf4 \strokec4 : \cf7 \strokec7 contCaducados\cf4 \strokec4  \};\cb1 \
\cb3 \}\cb1 \
\
\pard\pardeftab720\partightenfactor0
\cf5 \cb3 \strokec5 function\cf4 \strokec4  \cf7 \strokec7 enviarResumenSemanal\cf4 \strokec4 () \{\cb1 \
\pard\pardeftab720\partightenfactor0
\cf4 \cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 ss\cf4 \strokec4  = \cf6 \strokec6 SpreadsheetApp\cf4 \strokec4 .\cf7 \strokec7 openById\cf4 \strokec4 (\cf6 \strokec6 PROPS\cf4 \strokec4 .\cf7 \strokec7 getProperty\cf4 \strokec4 (\cf8 \strokec8 'DB_SS_ID'\cf4 \strokec4 ));\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 sheetConfig\cf4 \strokec4  = \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 getSheetByName\cf4 \strokec4 (\cf8 \strokec8 'CONFIG'\cf4 \strokec4 );\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 dataConfig\cf4 \strokec4  = \cf7 \strokec7 sheetConfig\cf4 \strokec4 .\cf7 \strokec7 getDataRange\cf4 \strokec4 ().\cf7 \strokec7 getValues\cf4 \strokec4 ();\cb1 \
\cb3   \cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 emailDestino\cf4 \strokec4  = \cf8 \strokec8 ""\cf4 \strokec4 ;\cb1 \
\cb3   \cf5 \strokec5 for\cf4 \strokec4 (\cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 row\cf4 \strokec4  \cf5 \strokec5 of\cf4 \strokec4  \cf7 \strokec7 dataConfig\cf4 \strokec4 ) \{ \cf5 \strokec5 if\cf4 \strokec4 (\cf7 \strokec7 row\cf4 \strokec4 [\cf9 \strokec9 0\cf4 \strokec4 ] === \cf8 \strokec8 'EMAIL_AVISOS'\cf4 \strokec4 ) \cf7 \strokec7 emailDestino\cf4 \strokec4  = \cf7 \strokec7 row\cf4 \strokec4 [\cf9 \strokec9 1\cf4 \strokec4 ]; \}\cb1 \
\cb3   \cf5 \strokec5 if\cf4 \strokec4 (!\cf7 \strokec7 emailDestino\cf4 \strokec4 ) \cf7 \strokec7 emailDestino\cf4 \strokec4  = \cf6 \strokec6 Session\cf4 \strokec4 .\cf7 \strokec7 getActiveUser\cf4 \strokec4 ().\cf7 \strokec7 getEmail\cf4 \strokec4 ();\cb1 \
\
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 hoy\cf4 \strokec4  = \cf5 \strokec5 new\cf4 \strokec4  \cf6 \strokec6 Date\cf4 \strokec4 ();\cb1 \
\cb3   \cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 alertas\cf4 \strokec4  = [];\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 sheetMant\cf4 \strokec4  = \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 getSheetByName\cf4 \strokec4 (\cf8 \strokec8 'PLAN_MANTENIMIENTO'\cf4 \strokec4 );\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 dataMant\cf4 \strokec4  = \cf7 \strokec7 sheetMant\cf4 \strokec4 .\cf7 \strokec7 getDataRange\cf4 \strokec4 ().\cf7 \strokec7 getValues\cf4 \strokec4 ();\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 mapActivos\cf4 \strokec4  = \{\};\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 dataAct\cf4 \strokec4  = \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 getSheetByName\cf4 \strokec4 (\cf8 \strokec8 'ACTIVOS'\cf4 \strokec4 ).\cf7 \strokec7 getDataRange\cf4 \strokec4 ().\cf7 \strokec7 getValues\cf4 \strokec4 ();\cb1 \
\cb3   \cf5 \strokec5 for\cf4 \strokec4 (\cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 i\cf4 \strokec4 =\cf9 \strokec9 1\cf4 \strokec4 ; \cf7 \strokec7 i\cf4 \strokec4 <\cf7 \strokec7 dataAct\cf4 \strokec4 .\cf7 \strokec7 length\cf4 \strokec4 ; \cf7 \strokec7 i\cf4 \strokec4 ++) \{ \cf7 \strokec7 mapActivos\cf4 \strokec4 [\cf7 \strokec7 dataAct\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 0\cf4 \strokec4 ]] = \cf7 \strokec7 dataAct\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 3\cf4 \strokec4 ]; \} \cb1 \
\
\cb3   \cf5 \strokec5 for\cf4 \strokec4 (\cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 i\cf4 \strokec4 =\cf9 \strokec9 1\cf4 \strokec4 ; \cf7 \strokec7 i\cf4 \strokec4 <\cf7 \strokec7 dataMant\cf4 \strokec4 .\cf7 \strokec7 length\cf4 \strokec4 ; \cf7 \strokec7 i\cf4 \strokec4 ++) \{\cb1 \
\cb3     \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 fecha\cf4 \strokec4  = \cf5 \strokec5 new\cf4 \strokec4  \cf6 \strokec6 Date\cf4 \strokec4 (\cf7 \strokec7 dataMant\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 4\cf4 \strokec4 ]);\cb1 \
\cb3     \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 diffDays\cf4 \strokec4  = \cf6 \strokec6 Math\cf4 \strokec4 .\cf7 \strokec7 ceil\cf4 \strokec4 ((\cf7 \strokec7 fecha\cf4 \strokec4  - \cf7 \strokec7 hoy\cf4 \strokec4 ) / (\cf9 \strokec9 1000\cf4 \strokec4  * \cf9 \strokec9 60\cf4 \strokec4  * \cf9 \strokec9 60\cf4 \strokec4  * \cf9 \strokec9 24\cf4 \strokec4 ));\cb1 \
\cb3     \cf5 \strokec5 if\cf4 \strokec4 (\cf7 \strokec7 diffDays\cf4 \strokec4  <= \cf9 \strokec9 30\cf4 \strokec4 ) \{\cb1 \
\cb3       \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 nombreActivo\cf4 \strokec4  = \cf7 \strokec7 mapActivos\cf4 \strokec4 [\cf7 \strokec7 dataMant\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 1\cf4 \strokec4 ]] || \cf8 \strokec8 "Desconocido"\cf4 \strokec4 ;\cb1 \
\cb3       \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 estado\cf4 \strokec4  = \cf7 \strokec7 diffDays\cf4 \strokec4  < \cf9 \strokec9 0\cf4 \strokec4  ? \cf8 \strokec8 "\uc0\u55357 \u56628  VENCIDO"\cf4 \strokec4  : \cf8 \strokec8 "\uc0\u55357 \u57313  PR\'d3XIMO"\cf4 \strokec4 ;\cb1 \
\cb3       \cf7 \strokec7 alertas\cf4 \strokec4 .\cf7 \strokec7 push\cf4 \strokec4 (\cf8 \strokec8 `<li><strong>\cf4 \strokec4 $\{\cf7 \strokec7 estado\cf4 \strokec4 \}\cf8 \strokec8 </strong>: \cf4 \strokec4 $\{\cf7 \strokec7 nombreActivo\cf4 \strokec4 \}\cf8 \strokec8  - \cf4 \strokec4 $\{\cf7 \strokec7 dataMant\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 2\cf4 \strokec4 ]\}\cf8 \strokec8  (\cf4 \strokec4 $\{\cf6 \strokec6 Utilities\cf4 \strokec4 .\cf7 \strokec7 formatDate\cf4 \strokec4 (\cf7 \strokec7 fecha\cf4 \strokec4 , \cf6 \strokec6 Session\cf4 \strokec4 .\cf7 \strokec7 getScriptTimeZone\cf4 \strokec4 (), \cf8 \strokec8 "dd/MM/yyyy"\cf4 \strokec4 )\}\cf8 \strokec8 )</li>`\cf4 \strokec4 );\cb1 \
\cb3     \}\cb1 \
\cb3   \}\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 sheetCont\cf4 \strokec4  = \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 getSheetByName\cf4 \strokec4 (\cf8 \strokec8 'CONTRATOS'\cf4 \strokec4 );\cb1 \
\cb3   \cf5 \strokec5 if\cf4 \strokec4 (\cf7 \strokec7 sheetCont\cf4 \strokec4 ) \{\cb1 \
\cb3     \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 dataCont\cf4 \strokec4  = \cf7 \strokec7 sheetCont\cf4 \strokec4 .\cf7 \strokec7 getDataRange\cf4 \strokec4 ().\cf7 \strokec7 getValues\cf4 \strokec4 ();\cb1 \
\cb3     \cf5 \strokec5 for\cf4 \strokec4 (\cf5 \strokec5 let\cf4 \strokec4  \cf7 \strokec7 i\cf4 \strokec4 =\cf9 \strokec9 1\cf4 \strokec4 ; \cf7 \strokec7 i\cf4 \strokec4 <\cf7 \strokec7 dataCont\cf4 \strokec4 .\cf7 \strokec7 length\cf4 \strokec4 ; \cf7 \strokec7 i\cf4 \strokec4 ++) \{\cb1 \
\cb3       \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 fecha\cf4 \strokec4  = \cf5 \strokec5 new\cf4 \strokec4  \cf6 \strokec6 Date\cf4 \strokec4 (\cf7 \strokec7 dataCont\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 6\cf4 \strokec4 ]);\cb1 \
\cb3       \cf5 \strokec5 if\cf4 \strokec4 (\cf7 \strokec7 fecha\cf4 \strokec4  < \cf7 \strokec7 hoy\cf4 \strokec4 ) \cf7 \strokec7 alertas\cf4 \strokec4 .\cf7 \strokec7 push\cf4 \strokec4 (\cf8 \strokec8 `<li><strong>\uc0\u55357 \u56628  CONTRATO CADUCADO</strong>: Proveedor \cf4 \strokec4 $\{\cf7 \strokec7 dataCont\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 3\cf4 \strokec4 ]\}\cf8 \strokec8  (Ref: \cf4 \strokec4 $\{\cf7 \strokec7 dataCont\cf4 \strokec4 [\cf7 \strokec7 i\cf4 \strokec4 ][\cf9 \strokec9 4\cf4 \strokec4 ]\}\cf8 \strokec8 )</li>`\cf4 \strokec4 );\cb1 \
\cb3     \}\cb1 \
\cb3   \}\cb1 \
\
\cb3   \cf5 \strokec5 if\cf4 \strokec4 (\cf7 \strokec7 alertas\cf4 \strokec4 .\cf7 \strokec7 length\cf4 \strokec4  > \cf9 \strokec9 0\cf4 \strokec4 ) \{\cb1 \
\cb3     \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 htmlBody\cf4 \strokec4  = \cf8 \strokec8 `<h3>Resumen Semanal GMAO</h3><ul>\cf4 \strokec4 $\{\cf7 \strokec7 alertas\cf4 \strokec4 .\cf7 \strokec7 join\cf4 \strokec4 (\cf8 \strokec8 ''\cf4 \strokec4 )\}\cf8 \strokec8 </ul>`\cf4 \strokec4 ;\cb1 \
\cb3     \cf6 \strokec6 MailApp\cf4 \strokec4 .\cf7 \strokec7 sendEmail\cf4 \strokec4 (\{ \cf7 \strokec7 to\cf4 \strokec4 : \cf7 \strokec7 emailDestino\cf4 \strokec4 , \cf7 \strokec7 subject\cf4 \strokec4 : \cf8 \strokec8 "\uc0\u9888 \u65039  Alerta GMAO: Vencimientos"\cf4 \strokec4 , \cf7 \strokec7 htmlBody\cf4 \strokec4 : \cf7 \strokec7 htmlBody\cf4 \strokec4  \});\cb1 \
\cb3   \}\cb1 \
\cb3 \}\cb1 \
\
\pard\pardeftab720\partightenfactor0
\cf2 \cb3 \strokec2 // ==========================================\cf4 \cb1 \strokec4 \
\cf2 \cb3 \strokec2 // 10. \'c1RBOL UI\cf4 \cb1 \strokec4 \
\cf2 \cb3 \strokec2 // ==========================================\cf4 \cb1 \strokec4 \
\pard\pardeftab720\partightenfactor0
\cf5 \cb3 \strokec5 function\cf4 \strokec4  \cf7 \strokec7 getArbolDatos\cf4 \strokec4 () \{\cb1 \
\pard\pardeftab720\partightenfactor0
\cf4 \cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 ss\cf4 \strokec4  = \cf6 \strokec6 SpreadsheetApp\cf4 \strokec4 .\cf7 \strokec7 openById\cf4 \strokec4 (\cf6 \strokec6 PROPS\cf4 \strokec4 .\cf7 \strokec7 getProperty\cf4 \strokec4 (\cf8 \strokec8 'DB_SS_ID'\cf4 \strokec4 ));\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 campus\cf4 \strokec4  = \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 getSheetByName\cf4 \strokec4 (\cf8 \strokec8 'CAMPUS'\cf4 \strokec4 ).\cf7 \strokec7 getDataRange\cf4 \strokec4 ().\cf7 \strokec7 getValues\cf4 \strokec4 ().\cf7 \strokec7 slice\cf4 \strokec4 (\cf9 \strokec9 1\cf4 \strokec4 );\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 edificios\cf4 \strokec4  = \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 getSheetByName\cf4 \strokec4 (\cf8 \strokec8 'EDIFICIOS'\cf4 \strokec4 ).\cf7 \strokec7 getDataRange\cf4 \strokec4 ().\cf7 \strokec7 getValues\cf4 \strokec4 ().\cf7 \strokec7 slice\cf4 \strokec4 (\cf9 \strokec9 1\cf4 \strokec4 );\cb1 \
\cb3   \cf5 \strokec5 const\cf4 \strokec4  \cf7 \strokec7 activos\cf4 \strokec4  = \cf7 \strokec7 ss\cf4 \strokec4 .\cf7 \strokec7 getSheetByName\cf4 \strokec4 (\cf8 \strokec8 'ACTIVOS'\cf4 \strokec4 ).\cf7 \strokec7 getDataRange\cf4 \strokec4 ().\cf7 \strokec7 getValues\cf4 \strokec4 ().\cf7 \strokec7 slice\cf4 \strokec4 (\cf9 \strokec9 1\cf4 \strokec4 );\cb1 \
\cb3   \cb1 \
\cb3   \cf5 \strokec5 return\cf4 \strokec4  \cf7 \strokec7 campus\cf4 \strokec4 .\cf7 \strokec7 map\cf4 \strokec4 (\cf7 \strokec7 c\cf4 \strokec4  => (\{\cb1 \
\cb3     \cf7 \strokec7 id\cf4 \strokec4 : \cf7 \strokec7 c\cf4 \strokec4 [\cf9 \strokec9 0\cf4 \strokec4 ], \cf7 \strokec7 nombre\cf4 \strokec4 : \cf7 \strokec7 c\cf4 \strokec4 [\cf9 \strokec9 1\cf4 \strokec4 ], \cf7 \strokec7 type\cf4 \strokec4 : \cf8 \strokec8 'CAMPUS'\cf4 \strokec4 ,\cb1 \
\cb3     \cf7 \strokec7 hijos\cf4 \strokec4 : \cf7 \strokec7 edificios\cf4 \strokec4 .\cf7 \strokec7 filter\cf4 \strokec4 (\cf7 \strokec7 e\cf4 \strokec4  => \cf6 \strokec6 String\cf4 \strokec4 (\cf7 \strokec7 e\cf4 \strokec4 [\cf9 \strokec9 1\cf4 \strokec4 ]) === \cf6 \strokec6 String\cf4 \strokec4 (\cf7 \strokec7 c\cf4 \strokec4 [\cf9 \strokec9 0\cf4 \strokec4 ])).\cf7 \strokec7 map\cf4 \strokec4 (\cf7 \strokec7 e\cf4 \strokec4  => (\{\cb1 \
\cb3       \cf7 \strokec7 id\cf4 \strokec4 : \cf7 \strokec7 e\cf4 \strokec4 [\cf9 \strokec9 0\cf4 \strokec4 ], \cf7 \strokec7 nombre\cf4 \strokec4 : \cf7 \strokec7 e\cf4 \strokec4 [\cf9 \strokec9 2\cf4 \strokec4 ], \cf7 \strokec7 type\cf4 \strokec4 : \cf8 \strokec8 'EDIFICIO'\cf4 \strokec4 ,\cb1 \
\cb3       \cf7 \strokec7 hijos\cf4 \strokec4 : \cf7 \strokec7 activos\cf4 \strokec4 .\cf7 \strokec7 filter\cf4 \strokec4 (\cf7 \strokec7 a\cf4 \strokec4  => \cf6 \strokec6 String\cf4 \strokec4 (\cf7 \strokec7 a\cf4 \strokec4 [\cf9 \strokec9 1\cf4 \strokec4 ]) === \cf6 \strokec6 String\cf4 \strokec4 (\cf7 \strokec7 e\cf4 \strokec4 [\cf9 \strokec9 0\cf4 \strokec4 ])).\cf7 \strokec7 map\cf4 \strokec4 (\cf7 \strokec7 a\cf4 \strokec4  => (\{\cb1 \
\cb3         \cf7 \strokec7 id\cf4 \strokec4 : \cf7 \strokec7 a\cf4 \strokec4 [\cf9 \strokec9 0\cf4 \strokec4 ], \cf7 \strokec7 nombre\cf4 \strokec4 : \cf7 \strokec7 a\cf4 \strokec4 [\cf9 \strokec9 3\cf4 \strokec4 ], \cf7 \strokec7 type\cf4 \strokec4 : \cf8 \strokec8 'ACTIVO'\cf4 \cb1 \strokec4 \
\cb3       \}))\cb1 \
\cb3     \}))\cb1 \
\cb3   \}));\cb1 \
\cb3 \}\cb1 \
}