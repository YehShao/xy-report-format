const Location = require('./Location.js');
module.exports = class NameListTable {
    constructor(sheet) {
        this.sheet = sheet
        this.leftBottom = new Location(sheet['!ref'].split(':')[1]);
        this.rightUpper = new Location('A1')
        this.getRightUpperConer()
    }
    inside(location) {
        return (!(location.after(this.leftBottom) || location.below(this.leftBottom) || this.rightUpper.below(location) || this.rightUpper.after(location)))
    }
    backToLeft(location) {
        location.x = this.rightUpper.x
    }
    backToTop(location) {
        location.y = this.rightUpper.y
    }
    nextByRow(location) {
        if (!this.inside(location)) {
            return false
        }
        if (this.leftBottom.equals(location)) {
            return false
        }
        if (location.x == this.leftBottom.x) {
            location.moveDown()
            this.backToLeft(location)
        } else {
            location.moveRight()
        }
        return true
    }
    nextByColumn(location) {
        if (!this.inside(location)) {
            return false
        }
        if (this.leftBottom.equals(location)) {
            return false
        }
        if (location.y == this.leftBottom.y) {
            location.moveRight()
            this.backToTop(location)
        } else {
            location.moveDown()
        }
        return true
    }
    valueOf(location) {
        if (!this.sheet[location.address()]) {
            return null
        }
        return this.sheet[location.address()].v
    }
    getRightUpperConer() {
        var parsePoint = new Location(this.rightUpper.address())        
        this.noLocations = []
        while (this.nextByColumn(parsePoint)) {
            if (this.sheet[parsePoint.address()] == null) {
                continue
            }
            if (this.sheet[parsePoint.address()].v == 'No.') {
                this.noLocations.push(new Location(parsePoint.address()))
            }
        }
        this.rightUpper = this.noLocations[0]
    }
    parseTable() {
        var nameRecordArray = []
        this.noLocations.forEach((noPoint) => {
            var pointer = new Location(noPoint.address())
            pointer.moveDown()
            while (this.valueOf(pointer)) {
                var num, name, att
                var subPointer = new Location(pointer.address())
                pointer.moveDown()
                num = parseInt(this.valueOf(subPointer))
                subPointer.moveRight()
                if (this.valueOf(subPointer)) {
                    name = this.valueOf(subPointer)
                } else {
                    continue
                }           
                subPointer.moveRight()
                att = this.valueOf(subPointer) == 'V'
                nameRecordArray.push({order: num, studentName: name, attendance: att})
            }
        })
        var result = nameRecordArray.filter(nameRecord => nameRecord.attendance)
        return result.sort(function(a, b) { return (a.order > b.order) ? 1 : ((b.order > a.order) ? -1 : 0)})
    }
}