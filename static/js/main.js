import Datepicker from "../../node_modules/flowbite-datepicker/Datepicker";
import ja from "../../node_modules/flowbite-datepicker/locales/ja";

const datepickerEl = document.getElementById("datepicker-actions");
Object.assign(Datepicker.locales, ja);
const datePicker = new Datepicker(datepickerEl, {
  language: 'ja',
});