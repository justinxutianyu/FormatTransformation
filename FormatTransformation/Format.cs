using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

/// <summary>
/// This file prepares vaious objects for the foramt transformation from Bibendum to Pricebook.
/// </summary>

namespace Format
{
    //these are for pricebook objects
    class Priceboo_Manufacture
    {
        //Constructor
        public Priceboo_Manufacture() {
        }

        private string id;
        private string name;
        private string description;
        private string manufacturerTemplateId;
        private string metaKeywords;
        private string metaDescription;
        private string metaTitle;
        private string seName;
        private string picture;
        private string pageSize;
        private string allowCustomersToSelectPageSize;
        private string pageSizeOptions;
        private string priceRanges;
        private string published;
        private string displayOrder;

        public string Id { get => Id; set => Id = value; }
        public string Name { get => name; set => name = value; }
        public string Description { get => description; set => description = value; }
        public string ManufacturerTemplateId { get => manufacturerTemplateId; set => manufacturerTemplateId = value; }
        public string MetaKeywords { get => metaKeywords; set => metaKeywords = value; }
        public string MetaDescription { get => metaDescription; set => metaDescription = value; }
        public string MetaTitle { get => metaTitle; set => metaTitle = value; }
        public string SeName { get => seName; set => seName = value; }
        public string Picture { get => picture; set => picture = value; }
        public string PageSize { get => pageSize; set => pageSize = value; }
        public string AllowCustomersToSelectPageSize { get => allowCustomersToSelectPageSize; set => allowCustomersToSelectPageSize = value; }
        public string PageSizeOptions { get => PageSizeOptions; set => PageSizeOptions = value; }
        public string PriceRanges { get => priceRanges; set => priceRanges = value; }
        public string Published { get => published; set => published = value; }
        public string DisplayOrder { get => displayOrder; set => displayOrder = value; }

        public void Initializer() {
           manufacturerTemplateId = "1";
           PageSize = "6";
           AllowCustomersToSelectPageSize = "TRUE";
           PageSizeOptions = "6,3,9";
           PriceRanges = "TRUE";
           DisplayOrder = "0";

        }
    }

    class Priceboo_Category {
        
        //Constructor
        public Priceboo_Category() {
        }

        private string id;
        private string name;
        private string description;
        private string categoryTemplateId;
        private string metaKeywords;
        private string metaDescription;
        private string metaTitle;
        private string seName;
        private string parentCategoryId;
        private string picture;
        private string pageSize;
        private string allowCustomersToSelectPageSize;
        private string pageSizeOptions;
        private string priceRanges;
        private string showOnHomePage;
        private string includeInTopMenu;
        private string published;
        private string displayOrder;

        public string Id { get => id; set => id = value; }
        public string Name { get => name; set => name = value; }
        public string Description { get => description; set => description = value; }
        public string CategoryTemplateId { get => categoryTemplateId; set => categoryTemplateId = value; }
        public string MetaKeywords { get => metaKeywords; set => metaKeywords = value; }
        public string MetaDescription { get => metaDescription; set => metaDescription = value; }
        public string MetaTitle { get => metaTitle; set => metaTitle = value; }
        public string SeName { get => seName; set => seName = value; }
        public string ParentCategoryId { get => parentCategoryId; set => parentCategoryId = value; }
        public string Picture { get => picture; set => picture = value; }
        public string PageSize { get => pageSize; set => pageSize = value; }
        public string AllowCustomersToSelectPageSize { get => allowCustomersToSelectPageSize; set => allowCustomersToSelectPageSize = value; }
        public string PageSizeOptions { get => pageSizeOptions; set => pageSizeOptions = value; }
        public string PriceRanges { get => priceRanges; set => priceRanges = value; }
        public string ShowOnHomePage { get => showOnHomePage; set => showOnHomePage = value; }
        public string IncludeInTopMenu { get => includeInTopMenu; set => includeInTopMenu = value; }
        public string Published { get => published; set => published = value; }
        public string DisplayOrder { get => displayOrder; set => displayOrder = value; }

        public void Initializer() {
            CategoryTemplateId = "1";
            PageSize = "6";
            AllowCustomersToSelectPageSize = "TRUE";
            PageSizeOptions = "6,3,9";
            ShowOnHomePage = "False";
            IncludeInTopMenu = "False";
            Published = "TRUE";
        }



    }

    class Priceboo_Product {
        //Constructor
        public Priceboo_Product() {
        }

        private string productType;
        private string parentGroupedProductId;
        private string visibleIndividually;
        private string name;
        private string shortDescription;
        private string fullDescription; 
        private string vendor;
        private string productTemplate;
        private string showOnHomePage;
        private string metaKeywords;
        private string metaDescription;
        private string metaTitle;
        private string seName;
        private string allowCustomerReviews;
        private string published;
        private string sKU;
        private string manufacturerPartNumber;
        private string gtin;
        private string isGiftCard;
        private string giftCardType;
        private string overriddenGiftCardAmount;
        private string requireOtherProducts;
        private string requiredProductIds;
        private string automaticallyAddRequiredProducts;
        private string isDownload;
        private string downloadId;
        private string unlimitedDownloads;
        private string maxNumberOfDownloads;
        private string downloadActivationType;
        private string hasSampleDownload;
        private string sampleDownloadId;
        private string hasUserAgreement;
        private string userAgreementText;
        private string isRecurring;
        private string recurringCycleLength;
        private string recurringCyclePeriod;
        private string recurringTotalCycles;
        private string isRental;
        private string rentalPriceLength;
        private string rentalPricePeriod;
        private string isShipEnabled;
        private string isFreeShipping;
        private string shipSeparately;
        private string additionalShippingCharge;
        private string deliveryDate;
        private string isTaxExempt;
        private string taxCategory;
        private string isTelecommunicationsOrBroadcastingOrElectronicServices;
        private string manageInventoryMethod;
        private string useMultipleWarehouses;
        private string warehouseId;
        private string stockQuantity;
        private string displayStockAvailability;
        private string displayStockQuantity;
        private string minStockQuantity;
        private string lowStockActivity;
        private string notifyAdminForQuantityBelow;
        private string backorderMode;
        private string allowBackInStockSubscriptions;
        private string orderMinimumQuantity;
        private string orderMaximumQuantity;
        private string allowedQuantities;
        private string allowAddingOnlyExistingAttributeCombinations;
        private string notReturnable;
        private string disableBuyButton;
        private string disableWishlistButton;
        private string availableForPreOrder;
        private string preOrderAvailabilityStartDateTimeUtc;
        private string callForPrice;
        private string price;
        private string oldPrice;
        private string productCost;
        private string specialPrice;
        private string specialPriceStartDateTimeUtc;
        private string specialPriceEndDateTimeUtc;
        private string customerEntersPrice;
        private string minimumCustomerEnteredPrice;
        private string maximumCustomerEnteredPrice;
        private string basepriceEnabled;
        private string basepriceAmount;
        private string basepriceUnit;
        private string basepriceBaseAmount;
        private string basepriceBaseUnit;
        private string markAsNew;
        private string markAsNewStartDateTimeUtc;
        private string markAsNewEndDateTimeUtc;
        private string weight;
        private string length;
        private string width;
        private string height;
        private string categories;
        private string manufacturers;
        private string picture1;
        private string picture2;
        private string picture3;


    }

    //these are for Bibendum objects
    class Bib_Manufacture {

        public Bib_Manufacture(string id, string name)
        {

            this.id = id;
            this.name = name;

        }

        private string id;

        public string Id { get => id; set => id = value; }

        private string name;

        public string Name { get => name; set => name = value; }


    }

    class Bib_Category {
        //Constructor
        public Bib_Category() {
        }

        private string parentcategory;
        private string name;

    }

    class Bib_Product {
        //Constructor
        public Bib_Product() {
        }

        private string productid;
        private string productcode;
        private string product;
        private string cleanurl;
        private string weight;
        private string list_price;
        private string descr;
        private string fulldescr;
        private string keywords;
        private string metakeywords;
        private string metadescription;
        private string avail;
        private string rating;
        private string forsale;
        private string shipping_freight;
        private string free_shipping;
        private string discount_avail;
        private string min_amount;
        private string length;
        private string width;
        private string height;
        private string low_avail_limit;
        private string free_tax;
        private string categoryid;
        private string category;
        private string membership;
        private string price;
        private string taxes;
        private string add_date;
        private string views_stats;
        private string sales_stats;
        private string del_stats;
        private string small_item;
        private string separate_box;
        private string items_per_box;
        private string region;
        private string manufactureid;
        private string manufacturer;


    }


}
