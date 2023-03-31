const puppeteer = require("puppeteer");
const XLSX = require("xlsx-js-style");

(async () => {

    async function amazonSearch(url, browser, i) {
        const page = await browser.newPage()

        await page.setViewport({ width: 1920, height: 1080 })
        await page.goto(url, { waitUntil: "domcontentloaded" })
        await page.waitForSelector(".s-card-container")
        if (i > 1) {
            await page.click(".s-pagination-next")
            await page.waitForSelector(".s-card-container")
        }

        const products = await page.evaluate(() => {
            const links = Array.from(document.querySelectorAll(".s-card-container"))
            if (!links.length) return [{ Name: "No results for search" }]
            return links.map(link => {
                if (link.querySelector(".puis-label-popover-default")) return
                if (link.querySelector(".a-offscreen")) {

                    const prodName = link.querySelector("h2 span.a-color-base")?.textContent
                    const splitName = prodName.split("(")

                    let prodRating = link.querySelector(".a-section.a-spacing-none.a-spacing-top-micro .a-size-base")?.textContent
                    if (prodRating.length !== 3) prodRating = ""

                    let prodReviews = link.querySelector(".a-size-base.s-underline-text")?.textContent
                    if (prodReviews?.includes("$")) {
                        prodReviews = ""
                    } else {
                        prodReviews = prodReviews.replace(/[()]/g, "")
                    }

                    let dataObject = {
                        Link: link.querySelector(".a-link-normal")?.href,
                        Name: splitName[0],
                        Current_Price: link.querySelector(".a-offscreen")?.textContent,
                        Orginal_Price: "",
                        Rating: prodRating,
                        Review_Count: prodReviews
                       
                    }
                    if (link.querySelector(".a-text-price")) dataObject.Orginal_Price = link.querySelector(".a-text-price .a-offscreen")?.textContent
                    return dataObject
                } else {
                    return {
                        Name: "N/A",
                    }
                }
            })
        })
        return products
    }

    async function memExpSearch(url, browser) {
        const page = await browser.newPage()
        await page.setViewport({ width: 1920, height: 1080 })
        await page.goto(url, { waitUntil: "domcontentloaded" })
        await page.select(".c-cact-filter-filters__display select", "120")
        await page.waitForSelector(".c-cact-filter-filters__display select")
        const products = await page.evaluate(() => {
            const links = Array.from(document.querySelectorAll(".c-shca-icon-item"))
            if (!links.length) return [{ Name: "No results for search" }]

            return links.map(link => {
                const prodName = link.querySelector(".c-shca-icon-item__body-name a")?.textContent.trim()
                const splitName = prodName.split("(")
                let prodRating = ""
                if (link.querySelector(".c-shca-review-stars__icon")) prodRating = "Rating Exists"

                let dataObject = {
                    Link: link.querySelector(".c-shca-icon-item__body-image a")?.href,
                    Name: splitName[0],
                    Current_Price: link.querySelector(".c-shca-icon-item__summary-list span")?.textContent.trim(),
                    Orginal_Price: "",
                    Rating: prodRating
                }

                if (link.querySelector(".c-shca-icon-item__summary-regular span")?.textContent !== dataObject.Current_Price) dataObject.Orginal_Price = link.querySelector(".c-shca-icon-item__summary-regular span")?.textContent

                return dataObject
            })
        })
        return products
    }

    async function neweggSearch(url, browser) {
        const page = await browser.newPage()
        await page.setViewport({ width: 1920, height: 1080 })
        await page.goto(url, { waitUntil: "domcontentloaded" })
        await page.select(".list-tool-view select", "96")
        await page.waitForSelector(".item-cells-wrap.border-cells")
        const products = await page.evaluate(() => {
            const links = Array.from(document.querySelectorAll(".item-cell"))
            if (!links.length) return [{ Name: "No results for search" }]

            return links.map(link => {
                const prodName = link.querySelector(".item-title")?.textContent.trim()
                const splitName = prodName.split("(")
                let dataObject = {
                    Link: link.querySelector(".item-rating")?.href || '',
                    Name: splitName[0],
                    Current_Price:  link.querySelector(".price-current")?.textContent.split(/\s/)[0],
                    Orginal_Price: link.querySelector(".price-was-data")?.textContent.split(/\s/)[0] || "",
                    Rating: link.querySelector(".rating.rating-5")?.getAttribute("aria-label") || "No rating",
                    Review_Count: link.querySelector(".item-rating-num")?.textContent.replace(/[()]/g, "")
                }

                return dataObject
            })
        })
        return products
    }

    async function runSearch(searchTerm, website, webShort, browser) {
        let searchResults = []
        const url = `${website}${encodeURIComponent(searchTerm)}`

        switch (webShort) {
            case "Amazon":
                for (let i = 1; i <= 2; i++) {
                    console.log("Searching Amazon page#:", `${i}`)
                    searchResults.push(...await amazonSearch(url, browser, i))
                }
                break
            case "MemoryExp":
                console.log("Searching MemoryExp")
                searchResults.push(...await memExpSearch(url, browser))
                break
            case "Newegg":
                console.log("Searching Newegg")
                searchResults.push(...await neweggSearch(url, browser))
                break
        }
        return searchResults
    }

    const searchTerm = ["rtx 4080", "rtx 4070ti"]


    const websites = ["https://www.amazon.ca/s?k=", "https://www.memoryexpress.com/Search/Products?Search=", "https://www.newegg.ca/p/pl?d="]
    const websiteShort = ["Amazon", "MemoryExp","Newegg"]

    try {
        const wb = XLSX.utils.book_new()
        const browser = await puppeteer.launch()

        for (let i = 0; i < searchTerm.length; i++) {
            console.log(`Search Query: ${searchTerm[i]}`)
            for (let j = 0; j < websites.length; j++) {
                const products = await runSearch(searchTerm[i], websites[j], websiteShort[j], browser)
                const sheet = XLSX.utils.json_to_sheet(products)
                const js = XLSX.utils.sheet_to_json(sheet, { header: 1 })

                js.forEach((row, index) => {
                    const urlIndex = 0
                    if (index > 0) {
                        const hyperlink = {
                            t: "s",
                            f: `=HYPERLINK("${row[urlIndex]}","${websiteShort[j]}")`
                        }
                        const font = {
                            color: { rgb: "0000FF" },
                            underline: true
                        }
                        row[urlIndex] = {
                            v: `"${websiteShort[j]}"`,
                            f: hyperlink.f,
                            t: hyperlink.t,
                            s: { font: font }
                        }
                    }
                })

                const sheetWithLinks = XLSX.utils.aoa_to_sheet(js)
                XLSX.utils.book_append_sheet(wb, sheetWithLinks, `${websiteShort[j]}${i + 1}`)
            }
        }
        await browser.close()
        XLSX.writeFile(wb, "Internet-Search.xlsx")

    } catch (e) {
        console.log("failed", e)
    }
})()


