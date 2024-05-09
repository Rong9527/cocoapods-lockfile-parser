const yaml = require('js-yaml');
const xlsx = require('node-xlsx');
const fs = require('fs');
const myConfig = require('./config')

let batch = 0 //层数，倒序
let xlsxSheet = [['层级', '组件名']] //excel数据
let range = []//单元格合并
let rangeOffet = 0

try {
    const doc = yaml.load(fs.readFileSync(myConfig.sourcePath, 'utf8'))

    const { PODS } = doc

    prune(PODS)

    writeExcel()

} catch (e) {
    console.log(e);
}

//生成依赖层级数据
function prune(tempPods) {
    if (!tempPods || tempPods.length === 0) {
        return
    }

    let leaves = []

    batch++

    //找到所有叶子节点
    tempPods.forEach(aPod => {
        if (typeof aPod == 'string') {
            const podName = aPod.split(' ')[0]
            if (leaves.indexOf(podName) != -1) {
                return
            }
            leaves.push(podName)
            xlsxSheet.push([`第${batch}层`, podName])
        }
        else if (typeof aPod == 'object') {
            const podName = Object.keys(aPod)[0]
            const podNameWithOutVersion = podName.split(' ')[0]
            const dependencies = aPod[podName]
            if (!dependencies || dependencies.length == 0) {
                if (leaves.indexOf(podNameWithOutVersion) != -1) {
                    return
                }
                leaves.push(podNameWithOutVersion)
                xlsxSheet.push([`第${batch}层`, podNameWithOutVersion])
            }
        }
    })

    //合并单元格
    const range0 = { s: { c: 0, r: rangeOffet+1 }, e: { c: 0, r: rangeOffet = rangeOffet+leaves.length } }
    range.push(range0)
  
    //console.log(`第${batch}层组件个数：${leaves.length}`)
    //console.log(leaves)
    //console.log(range0)

    //剪掉叶子
    const leftNodes = []
    tempPods.forEach(aPod => {
        if (typeof aPod == 'object') {
            const podName = Object.keys(aPod)[0]
            const dependencies = aPod[podName]
            if (dependencies && dependencies.length > 0) {
                const tempDependencies = []
                dependencies.forEach(dependency => {
                    if (leaves.indexOf(dependency.split(' ')[0]) == -1) {
                        tempDependencies.push(dependency)
                    }
                })
                aPod[podName] = tempDependencies
                leftNodes.push(aPod)
            }
        }

    })
    prune(leftNodes)
}
//生成excel
function writeExcel() {
    const excelData = [
        {
            name: 'sheet1',
            data: xlsxSheet
        }
    ]

    const options = { '!cols': [{ wch: 10 }, { wch: 80 }, { wch: 15 }], '!merges': range }
    //console.log(options)
    const buffer = xlsx.build(excelData, {sheetOptions:options})
    fs.writeFile('./resut.xlsx', buffer, function (err) {
        if (err) throw err
        console.log('写入到 xlsx 完成！')
    })
}

