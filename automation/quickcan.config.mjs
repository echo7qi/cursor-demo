import path from 'node:path';

export const QUICKCAN_CONFIG = {
  // 建议把浏览器绑定目录也指向这个路径，这样页面刷新后就会读取到自动更新的数据
  dataRoot:
    process.env.QUICKCAN_DATA_ROOT || path.resolve('/Users/wangqiqi/Desktop/dashboard-data'),
  authStatePath:
    process.env.QUICKCAN_AUTH_STATE || path.resolve('automation/.auth/quickcan.json'),
  timeoutMs: 45_000,
  boards: [
    {
      id: 'ops',
      name: '运营宣推',
      url: 'https://data-analytics.quickcan.com/page/w8d10552655344747b05e949',
      subdir: '运营宣推',
      filePrefix: '运营宣推_作品明细表_日',
      filters: [{ label: '时间维度', option: '日' }],
      clickTexts: ['作品明细表sheet', '作品明细表', '作品明细'],
    },
    {
      id: 'wishbar',
      name: '祈愿宣发',
      url: 'https://data-analytics.quickcan.com/page/j843e54458386404da258a0e',
      subdir: '祈愿',
      filePrefix: '祈愿用户来源贡献_周_日均',
      filters: [
        { label: '统计周期', option: '周' },
        { label: '周维度', option: '日均' },
      ],
      clickTexts: ['祈愿用户来源贡献'],
    },
    {
      id: 'reach',
      name: '目标用户触达',
      url: 'https://data-analytics.quickcan.com/page/u7815efbcf133446cad3b8fe',
      subdir: '触达',
      filePrefix: '目标用户触达_抽池',
      filters: [{ label: '项目类型', option: '抽池' }],
      clickTexts: [],
    },
  ],
};

