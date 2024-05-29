SUMMARY



WEEKLY PROCESS:
- Update the Thrasio Product file if needed.

- Export the weekly TikTok Ads report for LW. Place it in the Inputs folder.
https://ads.tiktok.com/i18n/perf/campaign?aadvid=7284256543762333698

- Download the All Orders file as CSV for LW. Place it in the Inputs folder.
https://seller-us.tiktok.com/order?selected_sort=6&tab=all

- Download the Affiliate Orders report as CSV for LW. Place it in the Inputs 
folder.
https://affiliate-us.tiktok.com/product/order?shop_region=US

- Dowload the Analytics - Live & Video - Video Details - Affiliate accts report
for LW. Place it in the Inputs folder.
https://seller-us.tiktok.com/compass/video-analytics/video-details?shop_region=US

- Download the Insense Transaction History for LW. Place it in the Inputs folder.
https://app.insense.pro/billing/history

CATALOG:
For Catalogging we are storing a variable with ASIN_TTSPID format that stores the
title, Parent TTSPID, [SKU ID] of all included in the parent, [Seller SKU].

We then create a dictionary with key NAME-ASIN, value is the above variable.
This means many NAME-ASIN keys will share the same value. 


HOW TO MAINTAIN:
- Adding new ASINs to the processor: 
    + Add the ASIN to the Catalog section, making sure to assign the exact Title 
    to the variable, and TT Seller SKU.
    + ASIN variable naming convention is 
    PARENT ASIN - Parent TTSPID = TTS Title, Parent TTSPID, [SKU ID], [Seller SKU]
    + Add the SKU ID as key to ACTIVE_PRODUCT_LIST dictionary, and ASIN as value.
    What this does is basically attach the description to the SKU ID. 

- Changing Platform Fee:
    + Adjust PLATFORM_FEE in the Constants section.

- Adding new financial metris:
    + Add new financial metrics constant for each new metric. Map it to the exact
    name in the excel sheet.

- Financial metrics maintnaince:
    + Need to make sure that if TikTok changes the Headers of their files, such
    new Header names be updated in the process_all_orders function.