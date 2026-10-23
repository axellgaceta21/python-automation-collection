// Phase 2A data only. Load as a classic script before a future viewer script.
// Paths resolve from the repository root. No assets are fetched here.
// Only sheets[] are viewer tabs; missing/blocked records are diagnostics.
// deploymentReady is advisory: it does not prevent direct access to public files.
globalThis.PROJECTS_CONFIG = {
  "schemaVersion": 1,
  "deploymentReady": false,
  "deploymentBlocker": null,
  "viewerDefaults": {
    "scrollOrigin": "top-left",
    "resetScrollOnWorksheetChange": true,
    "initialScale": 1,
    "preserveAspectRatio": true,
    "requiredZoomActions": [
      "zoom-in",
      "zoom-out",
      "reset-100-percent",
      "fit-to-width",
      "fullscreen-or-expanded"
    ]
  },
  "projects": [
    {
      "id": "business-lead-scraper",
      "title": "Business Website Lead Scraper",
      "raw": {
        "type": "excel",
        "status": "ready",
        "workbook": {
          "filename": "business_websites.xlsx",
          "path": "/business-lead-scraper/raw/business_websites.xlsx",
          "downloadAvailable": true
        },
        "defaultWorksheet": "Websites_Raw",
        "sheets": [
          {
            "worksheet": "Websites_Raw",
            "label": "Source Websites",
            "image": "/business-lead-scraper/raw/business-lead-scraper-raw.png",
            "width": 1757,
            "height": 481,
            "bytes": 186905,
            "aspectRatio": 3.652807,
            "shape": "wide",
            "coverageNote": "Screenshot preview; use the approved workbook download for full cell data."
          }
        ]
      },
      "output": {
        "type": "excel",
        "status": "ready",
        "workbook": {
          "filename": "business-leads-demo.xlsx",
          "path": "/business-lead-scraper/output/business-leads-demo.xlsx",
          "downloadAvailable": true
        },
        "defaultWorksheet": "Business_Leads",
        "sheets": [
          {
            "worksheet": "Business_Leads",
            "label": "Valid Leads",
            "image": "/business-lead-scraper/output/business-leads-demo.png",
            "width": 8112,
            "height": 1298,
            "bytes": 549842,
            "aspectRatio": 6.249615,
            "shape": "wide",
            "coverageNote": "All 13 synthetic demonstration records and 18 lead columns. Example domains are non-live; no real business contacts are shown."
          },
          {
            "worksheet": "Invalid_URLs",
            "label": "Invalid URLs",
            "image": "/business-lead-scraper/output/business-lead-invalid-url.png",
            "width": 1895,
            "height": 524,
            "bytes": 24304,
            "aspectRatio": 3.616412,
            "shape": "wide",
            "coverageNote": "Screenshot preview; use the approved workbook download for full cell data."
          }
        ],
        "downloadAvailable": true
      }
    },
    {
      "id": "clean-and-validate-excel",
      "title": "Clean and Validate Excel",
      "raw": {
        "type": "excel",
        "status": "partial",
        "workbook": {
          "filename": "large_messy_dataset.xlsx",
          "path": "/clean-and-validate-excel/raw/large_messy_dataset.xlsx",
          "downloadAvailable": true
        },
        "defaultWorksheet": "Orders_Raw",
        "sheets": [
          {
            "worksheet": "Orders_Raw",
            "label": "Raw Orders",
            "image": "/clean-and-validate-excel/raw/clean-and-validate-excel-raw.png",
            "width": 1120,
            "height": 3421,
            "bytes": 610562,
            "aspectRatio": 0.32739,
            "shape": "tall",
            "coverageNote": "Orders preview only; capture ends around row 130 of 1201 worksheet rows."
          }
        ],
        "missingWorksheets": [
          {
            "worksheet": "Customers_Raw",
            "status": "needs-review",
            "reason": "No verified PNG supplied; not a viewer tab."
          }
        ]
      },
      "output": {
        "type": "excel",
        "status": "ready",
        "workbook": {
          "filename": "cleaned_dataset_06.30_1.xlsx",
          "path": "/clean-and-validate-excel/output/cleaned_dataset_06.30_1.xlsx",
          "downloadAvailable": true
        },
        "defaultWorksheet": "Customers_Valid",
        "sheets": [
          {
            "worksheet": "Customers_Valid",
            "label": "Valid Customers",
            "image": "/clean-and-validate-excel/output/clean-customers-valid.png",
            "width": 1262,
            "height": 3254,
            "bytes": 928144,
            "aspectRatio": 0.38783,
            "shape": "tall",
            "coverageNote": "Excerpt ending around row 104 of 433 worksheet rows."
          },
          {
            "worksheet": "Customers_Invalid",
            "label": "Invalid Customers",
            "image": "/clean-and-validate-excel/output/clean-customers-invalid.png",
            "width": 1517,
            "height": 2647,
            "bytes": 904601,
            "aspectRatio": 0.573102,
            "shape": "tall",
            "coverageNote": "Screenshot preview; use the approved workbook download for full cell data."
          },
          {
            "worksheet": "Orders_Valid",
            "label": "Valid Orders",
            "image": "/clean-and-validate-excel/output/clean-orders-valid.png",
            "width": 1257,
            "height": 3254,
            "bytes": 526742,
            "aspectRatio": 0.386294,
            "shape": "tall",
            "coverageNote": "Excerpt ending around row 104 of 1156 worksheet rows."
          },
          {
            "worksheet": "Orders_Invalid",
            "label": "Invalid Orders",
            "image": "/clean-and-validate-excel/output/clean-orders-invalid.png",
            "width": 1404,
            "height": 1469,
            "bytes": 357525,
            "aspectRatio": 0.955752,
            "shape": "standard",
            "coverageNote": "Column A (Order ID) is absent from this capture."
          }
        ]
      }
    },
    {
      "id": "product-catalog-auditor",
      "title": "Product Catalog Auditor",
      "raw": {
        "type": "excel",
        "status": "ready",
        "workbook": {
          "filename": "products_catalog_messy.xlsx",
          "path": "/product-catalog-auditor/raw/products_catalog_messy.xlsx",
          "downloadAvailable": true
        },
        "defaultWorksheet": "Products_Raw",
        "sheets": [
          {
            "worksheet": "Products_Raw",
            "label": "Raw Catalog",
            "image": "/product-catalog-auditor/raw/product-catalog-raw.png",
            "width": 1773,
            "height": 4249,
            "bytes": 1460241,
            "aspectRatio": 0.417275,
            "shape": "wide-and-tall",
            "coverageNote": "Excerpt ending around row 121 of 196 worksheet rows."
          }
        ]
      },
      "output": {
        "type": "excel",
        "status": "ready",
        "workbook": {
          "filename": "product_catalog_audit_06.30_1.xlsx",
          "path": "/product-catalog-auditor/output/product_catalog_audit_06.30_1.xlsx",
          "downloadAvailable": true
        },
        "defaultWorksheet": "Products_Valid",
        "sheets": [
          {
            "worksheet": "Products_Valid",
            "label": "Valid Products",
            "image": "/product-catalog-auditor/output/cleaned-product-catalog-valid.png",
            "width": 1804,
            "height": 5292,
            "bytes": 3520788,
            "aspectRatio": 0.340892,
            "shape": "wide-and-tall",
            "coverageNote": "Screenshot preview; use the approved workbook download for full cell data."
          },
          {
            "worksheet": "Products_Invalid",
            "label": "Invalid Products",
            "image": "/product-catalog-auditor/output/cleaned-product-catalog-invalid.png",
            "width": 1892,
            "height": 533,
            "bytes": 245348,
            "aspectRatio": 3.549719,
            "shape": "wide",
            "coverageNote": "Screenshot preview; use the approved workbook download for full cell data."
          }
        ]
      }
    },
    {
      "id": "prestashop-validator",
      "title": "PrestaShop Validator",
      "raw": {
        "type": "excel",
        "status": "ready",
        "workbook": {
          "filename": "prestashop_export_messy.xlsx",
          "path": "/prestashop-validator/raw/prestashop_export_messy.xlsx",
          "downloadAvailable": true
        },
        "defaultWorksheet": "PrestaShop_Raw",
        "sheets": [
          {
            "worksheet": "PrestaShop_Raw",
            "label": "Raw Products",
            "image": "/prestashop-validator/raw/pretashop-raw.png",
            "width": 2529,
            "height": 5269,
            "bytes": 7040862,
            "aspectRatio": 0.479977,
            "shape": "wide-and-tall",
            "coverageNote": "Excerpt ending around row 150 of 217 worksheet rows."
          }
        ]
      },
      "output": {
        "type": "excel",
        "status": "ready",
        "workbook": {
          "filename": "prestashop_validation_06.30_1.xlsx",
          "path": "/prestashop-validator/output/prestashop_validation_06.30_1.xlsx",
          "downloadAvailable": true
        },
        "defaultWorksheet": "PrestaShop_Valid",
        "sheets": [
          {
            "worksheet": "PrestaShop_Valid",
            "label": "Valid Products",
            "image": "/prestashop-validator/output/cleaned-pretashop-valid.png",
            "width": 2154,
            "height": 2744,
            "bytes": 3797517,
            "aspectRatio": 0.784985,
            "shape": "wide-and-tall",
            "coverageNote": "Excerpt ending around row 100 of 201 worksheet rows."
          },
          {
            "worksheet": "PrestaShop_Invalid",
            "label": "Invalid Products",
            "image": "/prestashop-validator/output/cleaned-pretashop-invalid.png",
            "width": 2430,
            "height": 483,
            "bytes": 146989,
            "aspectRatio": 5.031056,
            "shape": "wide",
            "coverageNote": "Screenshot preview; use the approved workbook download for full cell data."
          }
        ]
      }
    },
    {
      "id": "sales-dashboard",
      "title": "Sales Dashboard Generator",
      "raw": {
        "type": "excel",
        "status": "ready",
        "workbook": {
          "filename": "sales_orders_messy.xlsx",
          "path": "/sales-dashboard/raw/sales_orders_messy.xlsx",
          "downloadAvailable": true
        },
        "defaultWorksheet": "Orders_Raw",
        "sheets": [
          {
            "worksheet": "Orders_Raw",
            "label": "Raw Orders",
            "image": "/sales-dashboard/raw/sales-orders-raw.png",
            "width": 1609,
            "height": 3537,
            "bytes": 934138,
            "aspectRatio": 0.454905,
            "shape": "wide-and-tall",
            "coverageNote": "Excerpt ending around row 200 of 757 worksheet rows."
          }
        ]
      },
      "output": {
        "type": "excel",
        "status": "ready",
        "workbook": {
          "filename": "sales_dashboard_06.30_1.xlsx",
          "path": "/sales-dashboard/output/sales_dashboard_06.30_1.xlsx",
          "downloadAvailable": true
        },
        "defaultWorksheet": "Dashboard",
        "sheets": [
          {
            "worksheet": "Dashboard",
            "label": "Dashboard",
            "image": "/sales-dashboard/output/cleaned-sales-dashboard.png",
            "width": 1323,
            "height": 811,
            "bytes": 77233,
            "aspectRatio": 1.631319,
            "shape": "standard",
            "coverageNote": "Screenshot preview; use the approved workbook download for full cell data."
          },
          {
            "worksheet": "Orders_Clean",
            "label": "Clean Orders",
            "image": "/sales-dashboard/output/cleaned-sales-order-valid.png",
            "width": 1581,
            "height": 11474,
            "bytes": 10975458,
            "aspectRatio": 0.13779,
            "shape": "tall",
            "coverageNote": "Capture ends around row 750 of 751 worksheet rows; visible column letters skip J/K. Coverage needs review."
          },
          {
            "worksheet": "Orders_Invalid",
            "label": "Invalid Orders",
            "image": "/sales-dashboard/output/cleaned-sales-order-invalid.png",
            "width": 2113,
            "height": 180,
            "bytes": 114204,
            "aspectRatio": 11.738889,
            "shape": "wide",
            "coverageNote": "Screenshot preview; use the approved workbook download for full cell data."
          },
          {
            "worksheet": "Monthly_Sales",
            "label": "Monthly Sales",
            "image": "/sales-dashboard/output/cleaned-monthly-sales.png",
            "width": 797,
            "height": 223,
            "bytes": 37146,
            "aspectRatio": 3.573991,
            "shape": "standard",
            "coverageNote": "Screenshot preview; use the approved workbook download for full cell data."
          },
          {
            "worksheet": "Category_Sales",
            "label": "Category Sales",
            "image": "/sales-dashboard/output/cleaned-category-sales.png",
            "width": 501,
            "height": 172,
            "bytes": 19974,
            "aspectRatio": 2.912791,
            "shape": "standard",
            "coverageNote": "Screenshot preview; use the approved workbook download for full cell data."
          },
          {
            "worksheet": "Channel_Sales",
            "label": "Channel Sales",
            "image": "/sales-dashboard/output/cleaned-channel-sales.png",
            "width": 501,
            "height": 173,
            "bytes": 19451,
            "aspectRatio": 2.895954,
            "shape": "standard",
            "coverageNote": "Screenshot preview; use the approved workbook download for full cell data."
          },
          {
            "worksheet": "Top_Products",
            "label": "Top Products",
            "image": "/sales-dashboard/output/cleaned-top-product-sales.png",
            "width": 666,
            "height": 261,
            "bytes": 43634,
            "aspectRatio": 2.551724,
            "shape": "standard",
            "coverageNote": "Screenshot preview; use the approved workbook download for full cell data."
          }
        ]
      }
    },
    {
      "id": "seo-auditor",
      "title": "SEO Audit Tool",
      "raw": {
        "type": "excel",
        "status": "ready",
        "workbook": {
          "filename": "seo_urls_sample.xlsx",
          "path": "/seo-auditor/raw/seo_urls_sample.xlsx",
          "downloadAvailable": true
        },
        "defaultWorksheet": "URLs_Raw",
        "sheets": [
          {
            "worksheet": "URLs_Raw",
            "label": "Source URLs",
            "image": "/seo-auditor/raw/urls-raw.png",
            "width": 1297,
            "height": 304,
            "bytes": 61154,
            "aspectRatio": 4.266447,
            "shape": "standard",
            "coverageNote": "Screenshot preview; use the approved workbook download for full cell data."
          }
        ]
      },
      "output": {
        "type": "excel",
        "status": "partial",
        "workbook": {
          "filename": "seo_audit_06.30_1.xlsx",
          "path": "/seo-auditor/output/seo_audit_06.30_1.xlsx",
          "downloadAvailable": true
        },
        "defaultWorksheet": "SEO_Audit",
        "sheets": [
          {
            "worksheet": "SEO_Audit",
            "label": "SEO Audit",
            "image": "/seo-auditor/output/cleaned-seo-audit.png",
            "width": 4225,
            "height": 388,
            "bytes": 313363,
            "aspectRatio": 10.889175,
            "shape": "wide",
            "coverageNote": "Screenshot preview; use the approved workbook download for full cell data."
          }
        ],
        "missingWorksheets": [
          {
            "worksheet": "URLs_Invalid",
            "status": "needs-review",
            "reason": "No verified PNG supplied; not a viewer tab."
          }
        ]
      }
    },
    {
      "id": "pdf-extractor",
      "title": "PDF Extractor",
      "raw": {
        "type": "pdf",
        "status": "ready",
        "workbook": null,
        "defaultWorksheet": null,
        "sheets": [],
        "documents": [
          {
            "filename": "invoice_001_dental.pdf",
            "label": "Dental Invoice",
            "path": "/pdf-extractor/raw/invoice_001_dental.pdf",
            "downloadAvailable": true
          },
          {
            "filename": "invoice_002_construction.pdf",
            "label": "Construction Invoice",
            "path": "/pdf-extractor/raw/invoice_002_construction.pdf",
            "downloadAvailable": true
          },
          {
            "filename": "invoice_missing_fields.pdf",
            "label": "Invoice with Missing Fields",
            "path": "/pdf-extractor/raw/invoice_missing_fields.pdf",
            "downloadAvailable": true
          },
          {
            "filename": "purchase_order_001.pdf",
            "label": "Purchase Order",
            "path": "/pdf-extractor/raw/purchase_order_001.pdf",
            "downloadAvailable": true
          }
        ],
        "defaultDocument": "invoice_001_dental.pdf"
      },
      "output": {
        "type": "excel",
        "status": "partial",
        "workbook": {
          "filename": "pdf_extraction_report_07.01_1.xlsx",
          "path": "/pdf-extractor/output/pdf_extraction_report_07.01_1.xlsx",
          "downloadAvailable": true
        },
        "defaultWorksheet": "Extracted_Data",
        "sheets": [
          {
            "worksheet": "Extracted_Data",
            "label": "Extracted Data",
            "image": "/pdf-extractor/output/clean-extracted-data.png",
            "width": 2432,
            "height": 252,
            "bytes": 87196,
            "aspectRatio": 9.650794,
            "shape": "wide",
            "coverageNote": "Screenshot preview; use the approved workbook download for full cell data."
          }
        ],
        "missingWorksheets": [
          {
            "worksheet": "Invalid_PDFs",
            "status": "needs-review",
            "reason": "No verified PNG supplied; not a viewer tab."
          }
        ]
      }
    }
  ]
};

