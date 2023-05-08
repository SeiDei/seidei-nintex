import { LitElement, html, css } from 'lit-element';
import XLSX from 'xlsx';

@customElement('excel-lookup')
class ExcelLookup extends LitElement {
  static get styles() {
    return css
      .error {
        color: red;
      }
    ;
  }

  static getMetaConfig(): Promise<NintexPlugin> | NintexPlugin {
    return {
      controlName: 'excel lookup',
      fallbackDisableSubmit: false,
      version: '1.0',
      properties: {
        who: {
        file: { type: File },
      sheet: { type: String },
      key: { type: String },
      value: { type: String },
      result: { type: String },
      error: { type: String },
        }
      }}
  };

  constructor() {
    super();
    this.file = null;
    this.sheet = '';
    this.key = '';
    this.value = '';
    this.result = '';
    this.error = '';
  }

  async lookup() {
    this.error = '';

    // Read the Excel file
    const fileReader = new FileReader();
    fileReader.onload = () => {
      const data = new Uint8Array(fileReader.result);
      const workbook = XLSX.read(data, { type: 'array' });

      // Get the sheet data
      const sheet = workbook.Sheets[this.sheet];
      if (!sheet) {
        this.error = `Sheet not found: ${this.sheet}`;
        return;
      }

      // Find the matching row
      const rows = XLSX.utils.sheet_to_json(sheet);
      const row = rows.find(r => r[this.key] === this.value);
      if (row) {
        this.result = row[this.key];
      } else {
        this.result = '';
        this.error = `No record found with ${this.key}=${this.value}`;
      }
    };
    fileReader.readAsArrayBuffer(this.file);
  }

  render() {
    return html`
      <div>
        <label for="file">Excel file:</label>
        <input type="file" id="file" name="file" @change="${this._onFileChange}">
        <br>
        <label for="sheet">Sheet name:</label>
        <input type="text" id="sheet" name="sheet" .value="${this.sheet}" @input="${this._onInput}">
        <br>
        <label for="key">Key:</label>
        <input type="text" id="key" name="key" .value="${this.key}" @input="${this._onInput}">
        <br>
        <label for="value">Value:</label>
        <input type="text" id="value" name="value" .value="${this.value}" @input="${this._onInput}">
        <br>
        <button @click="${this.lookup}" ?disabled="${!this.file || !this.sheet || !this.key || !this.value}">Lookup</button>
        <div class="result">${this.result}</div>
        <div class="error">${this.error}</div>
      </div>
    `;
  }

  _onInput(event) {
    this[event.target.name] = event.target.value;
  }

  _onFileChange(event) {
    this.file = event.target.files[0];
  }
}

//customElements.define('excel-lookup', ExcelLookup);


// registering the web component
const elementName = 'excel-lookup';
customElements.define(elementName, ExcelLookup);
