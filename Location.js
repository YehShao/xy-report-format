const numberOfA_Z = 26;
const asciiMagicNumber = 64
module.exports = class Location {
    constructor(address) {
        //var validate = /^\D+\d+$/.test(address)
        var resultRegex = /^(\D+)(\d+)$/.exec(address)
        this.x = Location.convertColumnCodeToInt(resultRegex[1]);
        this.y = parseInt(resultRegex[2]);
    }
    static convertColumnCodeToInt(string) {
        var stringArray = string.split('')
        stringArray.reverse()
        return stringArray.reduce((total, amount, index) => {
            return total + (amount.charCodeAt() - asciiMagicNumber) * Math.pow(numberOfA_Z, index)
        }, 0)
    }
    static convertIntToColumnCode(number, string = '') {
        if (number == 0) {
            return string
        }  else {
            return Location.convertIntToColumnCode(~~(number / numberOfA_Z), String.fromCharCode((number % numberOfA_Z) + asciiMagicNumber) + string)
        }
    } 
    moveRight() {
        this.x = this.x + 1;
    } 
    moveDown(){
        this.y = this.y + 1;
    }
    address() {
        return String.fromCharCode(this.x + 64) + this.y.toString();
    }
    after(anotherLocation) {
        return this.x > anotherLocation.x
    }
    below(anotherLocation) {
        return this.y > anotherLocation.y
    }
    equals(anotherLocation) {
        return (this.x == anotherLocation.x) && (this.y == anotherLocation.y)
    }
}
