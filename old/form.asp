<%@ Language=VBScript %><% Option Explicit

'-----------------------------------------------------------------------
'--- FileUp Simple Upload Sample
'---
'--- This sample demonstrates how to perform a single file upload
'--- with FileUp

'--- Copyright (c) 2006 SoftArtisans, Inc.
'--- Mail: info@softartisans.com   http://www.softartisans.com
'-----------------------------------------------------------------------
%>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>SoftArtisans FileUp Simple Upload Sample</title>
</head>

<body>
    <p align="center"><img src="/safileupsamples/fileupse.gif" alt="SoftArtisans FileUp" /></p>

    <h3 align="center">FileUp Simple Upload Sample</h3><!--

    Note: For any form uploading a file, the ENCTYPE="multipart/form-data" and
    METHOD="POST" attributes MUST be present

    -->

    <form action="../asp/SAFileUpSE.asp" enctype="MULTIPART/FORM-DATA" method="post">
        <table align="center" width="600">
            <tr>
                <td align="right" valign="top">Enter Filename:</td><!--
                Note: Notice this form element is of TYPE="FILE"
                -->

                <td align="left"><input type="FILE" name="Filedata" /><br />
                 <i>Click "Browse" to select a file to upload</i></td>
            </tr>

            <tr>
                <td align="right">&#160;</td>

                <td align="left"><input type="SUBMIT" name="SUB1" value="Upload File" /></td>
            </tr>

            <tr>
                <td colspan="2">
                    <hr noshade="noshade" />
                    <b>Note:</b> This sample demonstrates how to perform a simple single file upload with FileUp
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
