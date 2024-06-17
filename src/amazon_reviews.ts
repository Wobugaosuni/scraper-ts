/**
 * 根据亚马逊的商品链接，获取评论并填充到excel表格上
 */

import axios from 'axios';
import * as cheerio from 'cheerio';
import { v2 } from '@google-cloud/translate';
import ExcelJS from 'exceljs';
import * as fs from 'fs';

// 创建 Google Cloud Translate 客户端
const translate = new v2.Translate({ key: 'AIzaSyCejislUsBfUgkszeaP7L5GMtu5rhpEF2Q' });

// 定义目标商品的亚马逊页面 URL
const url = 'https://www.amazon.com/MEEPO-Electric-Skateboard-32Mph-Motors/dp/B0C5QCVTBJ/ref=cm_cr_arp_d_product_top';  // 示例商品 URL

interface Review {
  englishContent: string;
  chineseContent: string;
  imageUrl: string;
  brand: string;
  source: string;
  userName: string;
  userLink: string;
  reviewLink: string;
}

// 获取页面内容
async function fetchPageContent(url: string) {
  const { data } = await axios.get(url, {
    headers: {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
  });
  return data;
}

// 解析评论内容
async function parseReviews(html: string): Promise<Review[]> {
  const $ = cheerio.load(html);
  const reviews: Review[] = [];

  $('.review').each((_, element) => {
    const englishContent = $(element).find('.review-text-content').text().trim();
    const chineseContent = '';  // Will be filled later
    const userName = $(element).find('.a-profile-name').text().trim();
    const userLink = 'https://www.amazon.com' + $(element).find('.a-profile').attr('href');
    const reviewLink = 'https://www.amazon.com' + $(element).find('.a-link-normal').attr('href');
    const imageUrl = $(element).find('.review-image-tile').attr('src') || '';
    console.log('imageUrl:', imageUrl)
    const source = '亚马逊';
    const brand = '';

    reviews.push({
      englishContent,
      chineseContent,
      imageUrl,
      brand,
      source,
      userName,
      userLink,
      reviewLink
    });
  });

  // Translate all reviews to Chinese
  for (const review of reviews) {
    const [translation] = await translate.translate(review.englishContent, 'zh-CN');
    review.chineseContent = translation;
  }

  return reviews;
}

// 创建 Excel 文件
async function createExcel(reviews: Review[]) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Amazon Reviews');

  // 添加表头
  worksheet.columns = [
    { header: '评论内容（英文）', key: 'englishContent', width: 30 },
    { header: '评论内容（中文）', key: 'chineseContent', width: 30 },
    { header: '评论图片', key: 'imageUrl', width: 15 },
    { header: '涉及品牌', key: 'brand', width: 15 },
    { header: '评论来源', key: 'source', width: 15 },
    // { header: '评论用户', key: 'userName', width: 15 },
    { header: '评论用户', key: 'userLink', width: 30 },
    { header: '评论链接', key: 'reviewLink', width: 30 }
  ];

  // 填充数据
  for (const review of reviews) {
    worksheet.addRow({
      englishContent: review.englishContent,
      chineseContent: review.chineseContent,
      imageUrl: review.imageUrl,
      brand: review.brand,
      source: review.source,
      // userName: review.userName,
      userLink: review.userLink,
      reviewLink: review.reviewLink
    });
  }

  // 插入图片
  for (let i = 0; i < reviews.length; i++) {
    const review = reviews[i];
    if (review.imageUrl) {
      const response = await axios.get(review.imageUrl, { responseType: 'arraybuffer' });
      const imageBuffer = Buffer.from(response.data, 'binary');
      const imageId = workbook.addImage({
        buffer: imageBuffer,
        extension: 'png',
      });
      worksheet.addImage(imageId, `C${i + 2}:C${i + 2}`);
    }
  }

  // 保存文件
  await workbook.xlsx.writeFile('amazon_reviews.xlsx');
}

// 主函数
async function main() {
  try {
    const html = await fetchPageContent(url);
    const reviews = await parseReviews(html);
    await createExcel(reviews);
    console.log('评价数据已保存到 amazon_reviews1.xlsx');
  } catch (error) {
    console.error('Error:', error);
  }
}

main();
