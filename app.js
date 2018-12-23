const axios = require('axios')
const qs = require('qs')
const node_xj = require("xls-to-json")
const result = {}
const report = []
const formatData = (data) => {
    const config = {
        url: data['url'],
        method: data['请求类型'],
    }
    if (data['header']) {
        config['headers'] = data['header']
    }
    if (data['请求数据']) {
        config['data'] = qs.stringify(JSON.parse(data['请求数据']))
    }
    if ('' !== data['case依赖']) {
        if (!config['headers']) {
            config['headers'] = {}
        }
        config['headers']['authorization'] = report.filter(d => d['id'] + '' === data['case依赖'])[0]['responseData'][data['数据依赖字段']]
    }
    return config
}
const sendHttp = (data, r) => new Promise((resolve, reject) => {
    axios(formatData(data)).then(res => {
        const resData = res.data
        const d = {
            id: data['id'],
            responseData: resData,
            result: eval(data['预期结果']),
            case: data
        }
        r.push(d)     
        resolve(d)
    }).catch((err) => {
        const d = {
            id: data['id'],
            error: err,
            result: false,
            case: data
        }
        r.push(d)
        resolve(d)
    })
})
const sendReq = (res) => {
    const noDependenceCase = []
    const hasDependenceCase = []
    res.filter(v => '' === v['case依赖']).map((r, i) => {
        noDependenceCase[i] = sendHttp(r, report)
    })
    Promise.all(noDependenceCase).then(() => {
        res.filter(v => '' !== v['case依赖']).map((r, i) => {
            hasDependenceCase[i] = sendHttp(r, report)
        })
        Promise.all(hasDependenceCase).then(() => {
            result['success'] = report.filter(d => d.result).length
            result['failed'] = report.filter(d => !d.result).length
            const finalReport = {
                total: result,
                report: report.sort((a, b) => a['id'] - b['id']),
            }
            console.log(JSON.stringify(finalReport))

        })
    })
}
node_xj({
  input: "case/case.xls",
  output: "case.json",
}, (err, res) => {
  if (err) {
    console.error(err)
  } else {
    result['totalCaseNumber'] = res.length
    // 过滤掉不运行的请求
    const testCase = res.filter(d => 'yes' === d['是否运行'])
    result['testCaseNumber'] = testCase.length
    sendReq(testCase)
  }
})