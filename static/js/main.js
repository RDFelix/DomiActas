// Si estás usando módulos ES6 (con 'import')
import { toWords } from 'number-to-words';

// Si estás usando CommonJS (en Node.js, con 'require')
// const { toWords } = require('number-to-words');

console.log(toWords(123));          // "one hundred and twenty-three"
console.log(toWords(7));            // "seven"
console.log(toWords(15));           // "fifteen"
console.log(toWords(42));           // "forty-two"
console.log(toWords(100));          // "one hundred"
console.log(toWords(1000));         // "one thousand"
console.log(toWords(12345));        // "twelve thousand three hundred and forty-five"
console.log(toWords(0));            // "zero"
console.log(toWords(-5));           // "minus five"
console.log(toWords(1000000));      // "one million"
console.log(toWords(123456789));    // "one hundred and twenty-three million four hundred and fifty-six thousand seven hundred and eighty-nine"