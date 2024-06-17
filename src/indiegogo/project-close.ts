import puppeteer from 'puppeteer';
import ExcelJS from 'exceljs';
import * as fs from 'fs';
import * as path from 'path';


// 搜索关键词：
// E-Board
// electric skateboard
const urlList = [
    'https://www.indiegogo.com/projects/yiiboard-the-coolest-electric-skateboard-ever/x/37902132#/',
    'https://www.indiegogo.com/projects/hurricanex-the-most-powerful-electric-skateboard/x/37902132#/',
    'https://www.indiegogo.com/projects/charge-boards-electric-skateboards-under-500/x/37902132#/',
    'https://www.indiegogo.com/projects/panzerboard-powerful-all-terrain-e-skateboard/x/37902132#/'
];

interface ProjectInfo {
    name: string | undefined;
    description: string | undefined;
    link: string;
    creator: string | undefined;
    location: string | undefined;
    startDate: string;
    endDate: string;
    duration: string | undefined;
    isSuccess: string;
    targetAmount: string | undefined;
    actualAmount: string | undefined;
    actualAmountUSD: string | undefined;
    backers: string | undefined;
}


const scrapeData = async (url: string): Promise<ProjectInfo> => {
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    
    try {
        await page.goto(url, { waitUntil: 'networkidle2', timeout: 60000 });
    } catch (error) {
        console.error('Navigation timeout:', error);
        await browser.close();
        process.exit(1);
    }

    const projectInfo = await page.evaluate(() => {
        const getText = (selector: string) => {
            const element = document.querySelector(selector);
            return element ? element.textContent?.trim?.() : '';
        };

        const name = getText('.basicsSection-title');
        const description = getText('div.basicsSection-tagline');
        const link = window.location.href;
        const creator = getText('.campaignOwnerName-tooltip');
        const location = getText('div.basicsCampaignOwner-details-city');

        const datesText = getText('div.campaignHeaderBasics-deadline');
        const dates = datesText ? datesText.split('–') : ['', ''];
        const startDate = dates.length > 0 ? dates[0].trim() : '';

        const endDateText = getText('.basicsGoalProgress-progressDetails-detailsGoal-goalMetDate')
        const endDate = endDateText ? endDateText.trim().split('on')[1] : '';

        const duration = getText('div.campaignHeaderBasics-duration');
        // 未知
        const isSuccess = '/'

        const targetAmount = getText('div.campaignHeaderBasics-goal span');
        const actualAmount = getText('.basicsGoalProgress-amountSold.t-rebrand-h4s');
        // const actualAmount = '';
        const actualAmountUSD = actualAmount; // Assuming the amounts are already in USD

        // 众筹人数
        const backers = getText('.basicsGoalProgress-claimedOrBackers span:nth-of-type(1)');
        // const backers = backersText ? backersText.split(' ')[0] : '';

        return {
            name,
            description,
            link,
            creator,
            location,
            startDate,
            endDate,
            duration,
            isSuccess,
            targetAmount,
            actualAmount,
            actualAmountUSD,
            backers,
        };
    });

    await browser.close();
    return projectInfo;
};

const createExcel = async (projectInfos: ProjectInfo[]) => {
    const filePath = path.join(__dirname, 'IndiegogoProjectInfo.xlsx');

    // 如果文件存在，先删除文件
    if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Indiegogo Project');

    worksheet.columns = [
        { header: '项目名称', key: 'name', width: 30 },
        { header: '项目描述', key: 'description', width: 50 },
        { header: '项目链接', key: 'link', width: 50 },
        { header: '项目发起人', key: 'creator', width: 30 },
        { header: '项目所在地区', key: 'location', width: 30 },
        { header: '众筹开始时间', key: 'startDate', width: 20 },
        { header: '众筹结束时间', key: 'endDate', width: 20 },
        { header: '众筹周期时长', key: 'duration', width: 20 },
        { header: '众筹是否成功', key: 'isSuccess', width: 15 },
        { header: '目标众筹金额', key: 'targetAmount', width: 20 },
        { header: '实际众筹金额', key: 'actualAmount', width: 20 },
        { header: '实际众筹金额（美元）', key: 'actualAmountUSD', width: 25 },
        { header: '实际众筹人数', key: 'backers', width: 15 },
    ];

    projectInfos.forEach((projectInfo) => {
        worksheet.addRow(projectInfo);
    });

    await workbook.xlsx.writeFile(filePath);

    await workbook.xlsx.writeFile('IndiegogoProjectInfo.xlsx');
    console.log('Excel file created: IndiegogoProjectInfo.xlsx');
};

const main = async () => {
    try {
        const projectInfos: ProjectInfo[] = [];
        for (const url of urlList) {
            const projectInfo = await scrapeData(url);
            projectInfos.push(projectInfo);
        }
        await createExcel(projectInfos);
    } catch (error) {
        console.error('Error:', error);
    }
};

main();
