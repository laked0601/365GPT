# 365GPT
  <h2>📌 Description</h2>
  <p>Integration of ChatGPT within Excel and planned for other Office 365 apps. This product is an independent development and is not endorsed or affiliated with Microsoft or OpenAI.</p>
  <p>365GPT integrates ChatGPT into Excel using the OpenAI REST API and ChatCompletion endpoint. It has many use cases, but may best be used for generating repeated queries from ChatGPT in a tabulated data format, hence being used in Excel.</p>
  <h2>📁 Installation</h2>
  <ul>
    <li><strong>Installing the VBA Module</strong></li>
  </ul>
  <ol>
    <li>Create a new Excel Workbook and save it as an Excel Macro Enabled (.xlsm) file.</li>
    <li>Download "OpenAI.bas".</li>
    <li>On the first sheet you have open in the workbook, right-click, and then select "View Code".</li>
    <li>In the new window that opens, click File, then Import File.</li>
    <li>Navigate to where you downloaded OpenAI.bas, select it, then click Open.</li>
    <li>Save your workbook - this completes the installation.</li>
  </ol>
  <ul>
    <li><strong>Getting Your OpenAI Key</strong></li>
  </ul>
  <ol>
    <li>Navigate to <a href="https://platform.openai.com/" target="_blank">https://platform.openai.com/</a>.</li>
    <li>Create your account, or sign in to your existing one.</li>
    <li>In the top right, click on your profile, then select "View API keys".</li>
    <li>This page lists your existing API keys. If you have a key you want to use, select it from this dropdown. Otherwise, select "Create new secret key" and copy the value. It should look like "sk-.....".
     <img src='https://github.com/laked0601/365GPT/assets/90655952/00b64d49-0bb2-47fb-99c4-776f87a52c4a'>
    </li>
    <li>Go into the (.xlsm) file you created earlier, in any empty cell type "=SetOpenAIKey(*YOUR_API_KEY*)" with the secret key you just copied.</li>
    <li>To test if this was assigned correctly, in a new empty cell, type "=GetOpenAIKey()". The resulting value should be the key you saved.</li>
    <li>In the case this doesn't work, you will need to include the key in all of your queries as the second parameter. For example: "=ChatCompletion('What President is on the $1 bill?', *YOUR_API_KEY*)".</li>
  </ol>
  <h2>⚠️ Known Limitations</h2>
  <ul>
    <li>Each query made with the function <code>=ChatCompletion()</code> uses OpenAI tokens, which are chargeable at the rates set by OpenAI for GPT-3.5 Turbo - 4k context. For more information, please read <a href="https://openai.com/pricing" target="_blank">https://openai.com/pricing</a>. As such, it is recommended to take actions like disabling auto-function renewal, pasting function output as text once completed, and limiting excessive function calls.</li>
    <li>Certain non-standard characters may not be parsed correctly as JSON when running a query. It is recommended that all new line and tab characters (CHAR(10) and CHAR(9)) are converted to their JSON safe equivalents ('\n', '\t') beforehand.</li>
  </ul>
