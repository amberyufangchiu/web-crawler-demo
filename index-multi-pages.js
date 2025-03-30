const puppeteer = require("puppeteer");
const ExcelJS = require("exceljs");

(async () => {
  const websiteLink = "WEBSITE_LINK_HERE";
  const browser = await puppeteer.launch({
    headless: false,
    slowMo: 100,
    ignoreHTTPSErrors: true,
    defaultViewport: null,
    ignoreDefaultArgs: ["--enable-automation"],
    args: [
      "--disable-infobars",
      "--no-sandbox",
      "--disable-setuid-sandbox",
      "--disable-gpu=False",
      "--enable-webgl",
      "--window-size=1600,900",
      "--start-maximized",
      // Args for AWS Lambda:
      "--single-process",
      "--user-data-dir=/tmp/user-data",
      "--data-path=/tmp/data-path",
      "--homedir=/tmp",
      "--disk-cache-dir=/tmp/cache-dir",
      "--database=/tmp/database",
    ],
  });
  const page = await browser.newPage();

  await page.goto(websiteLink, {
    waitUntil: "domcontentloaded",
    timeout: 120000,
  });

  await page.waitForSelector(".pagination");

  const totalPages = await page.evaluate(() => {
    const pages = document
      .querySelector(".pagination li:nth-last-child(2) a")
      ?.getAttribute("data-page");

    return parseInt(pages);
  });

  let allUrls = [];

  for (let i = 1; i <= totalPages; i++) {
    const categoryPage = await browser.newPage();
    await categoryPage.goto(`${websiteLink}?page=${i}`, {
      waitUntil: "domcontentloaded",
      timeout: 120000,
    });
    await categoryPage.waitForSelector(".product.product--teaser");

    const productUrls = await categoryPage.evaluate(() => {
      return Array.from(
        document.querySelectorAll(".product.product--teaser")
      ).map((el) => el.getAttribute("data-product-url"));
    });

    allUrls.push(...productUrls);

    await categoryPage.close();
  }

  let productsData = [];

  for (const url of allUrls) {
    const productPage = await browser.newPage();
    await productPage.goto(url, {
      waitUntil: "domcontentloaded",
      timeout: 120000,
    });

    const productData = await productPage.evaluate(() => {
      const title =
        document.querySelector(".product__title")?.innerText || "N/A";
      const price =
        document.querySelector(".product__price-value")?.innerText || "N/A";

      const sizes = Array.from(
        document.querySelectorAll(".product__option-values label")
      ).map((el) => el.innerText.trim());

      const descriptions = document.querySelectorAll(".description");

      let composition = "";
      let descriptionText = "";

      descriptions.forEach((description) => {
        const compositionParagraph = Array.from(
          description.querySelectorAll("p")
        ).find((p) => p.textContent.includes("Composition :"));

        if (compositionParagraph) {
          const compositionText = compositionParagraph.textContent
            .replace("Composition : ", "")
            .trim();
          composition = compositionText;

          const elementsBeforeComposition = Array.from(
            description.children
          ).slice(
            0,
            Array.from(description.children).indexOf(compositionParagraph)
          );

          const textBeforeComposition = elementsBeforeComposition
            .map((element) => {
              if (element.tagName === "P") {
                return element.textContent.trim();
              } else if (element.tagName === "UL") {
                return Array.from(element.querySelectorAll("li"))
                  .map((li) => li.textContent.trim())
                  .join(" ");
              }
            })
            .join(" ");

          descriptionText = textBeforeComposition;
        }
      });

      return { title, price, sizes, descriptionText, composition };
    });

    productsData.push({ url, ...productData });

    await productPage.close();
  }

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Products");

  worksheet.columns = [
    { header: "Title", key: "title", width: 40 },
    { header: "Price", key: "price", width: 20 },
    { header: "Sizes", key: "sizes", width: 30 },
    { header: "Description", key: "descriptionText", width: 80 },
    { header: "Composition", key: "composition", width: 30 },
    { header: "URL", key: "url", width: 100 },
  ];

  productsData.forEach((item) => {
    worksheet.addRow({
      title: item.title,
      price: item.price,
      sizes: item.sizes.join(", "),
      descriptionText: item.descriptionText,
      composition: item.composition,
      url: item.url,
    });
  });

  await workbook.xlsx.writeFile("products.xlsx");
  console.log("Excel file has been generated!");
})();
