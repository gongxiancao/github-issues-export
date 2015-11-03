var argv = require('optimist').argv;
var request = require('request');
var _ = require('lodash');
var excel = require('excel-export');
var fs = require('fs');
var moment = require('moment');

argv.out = argv.out || './issues.xlsx';

getIssues(1, 100, [], function (err, data) {
  data = prepareForExporting(data);
  exportToExcel(data, argv.out);
});

function getIssues (page, perPage, data, done) {
  console.log('getting response for page=' + page);
  request.get(
  {
    headers: {'user-agent': 'node.js'},
    url: 'https://api.github.com/repos/cloudnapps/PunchCard/issues?state=open&page=' + page + '&per_page=' + perPage,
    json: true,
    auth: {
      username: argv.username,
      password: argv.password
    }
  },
  function (err, res, body) {
    console.log('got response, page=' + page + ', size=' + body.length);
    if(res.statusCode === 200) {
      // console.log(page, data);
      Array.prototype.push.apply(data, body);
      if(body.length >= perPage) {
        setTimeout(getIssues.bind(null, page + 1, perPage, data, done), 1000);
      } else {
        done(null, data);
      }
    } else {
      done(new Error('error! status= ' + rs.statusCode + ', body=' + body));
    }
  });
}

function prepareForExporting (data) {
  data = _.map(data, function (issue) {
    return {
      number: issue.number,
      title: issue.title,
      asignee: (issue.assignee || {}).login ||'',
      labels: _.map(issue.labels, 'name').join(','),
      state: issue.state || '',
      milestone: (issue.milestone || {}).title ||'',
      due_on: moment((issue.milestone || {}).due_on || (new Date(2017, 0, 0).toISOString())).format('YYYY/MM/DD')
    };
  });

  data = _.sortBy(data, 'due_on');
  return data;
}

function exportToExcel (data, file) {
  var conf ={};
  // conf.stylesXmlFile = 'styles.xml';
  conf.cols = [
  {
    caption:'number',
    type:'number',
    width:28.7109375
  },
  {
    caption:'title',
    type:'string',
    width:28.7109375
  },
  {
    caption:'asignee',
    type:'string',
    width:28.7109375
  },
  {
    caption:'labels',
    type:'string',
    width:28.7109375
  },
  {
    caption:'state',
    type:'string',
    width:28.7109375
  },
  {
    caption:'milestone',
    type:'string',
    width:28.7109375
  },
  {
    caption:'due_on',
    type:'string',
    width:28.7109375
  }];

  conf.rows = _.map(data, function (item) {
    var row = _.map(conf.cols, function (col) {
      return item[col.caption];
    });
    return row;
  });

  var result = excel.execute(conf);

  var stream = fs.createWriteStream(file);

  result.pipe(stream);
}
