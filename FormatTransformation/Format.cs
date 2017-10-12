using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.Util;

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

        public string Id { get => id; set => id = value; }
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
        public string PageSizeOptions { get => pageSizeOptions; set => pageSizeOptions = value; }
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

        //specification attributes
        private string style;
        private string country;
        private string region;
        private string ingredients;
        private string attribute;
        private string bottlesize;
        private string containertype;
        private string vintage;


       

        public string ProductType { get => productType; set => productType = value; }
        public string ParentGroupedProductId { get => parentGroupedProductId; set => parentGroupedProductId = value; }
        public string VisibleIndividually { get => visibleIndividually; set => visibleIndividually = value; }
        public string Name { get => name; set => name = value; }
        public string ShortDescription { get => shortDescription; set => shortDescription = value; }
        public string FullDescription { get => fullDescription; set => fullDescription = value; }
        public string Vendor { get => vendor; set => vendor = value; }
        public string ProductTemplate { get => productTemplate; set => productTemplate = value; }
        public string ShowOnHomePage { get => showOnHomePage; set => showOnHomePage = value; }
        public string MetaKeywords { get => metaKeywords; set => metaKeywords = value; }
        public string MetaDescription { get => metaDescription; set => metaDescription = value; }
        public string MetaTitle { get => metaTitle; set => metaTitle = value; }
        public string SeName { get => seName; set => seName = value; }
        public string AllowCustomerReviews { get => allowCustomerReviews; set => allowCustomerReviews = value; }
        public string Published { get => published; set => published = value; }
        public string SKU { get => sKU; set => sKU = value; }
        public string ManufacturerPartNumber { get => manufacturerPartNumber; set => manufacturerPartNumber = value; }
        public string Gtin { get => gtin; set => gtin = value; }
        public string IsGiftCard { get => isGiftCard; set => isGiftCard = value; }
        public string GiftCardType { get => giftCardType; set => giftCardType = value; }
        public string OverriddenGiftCardAmount { get => overriddenGiftCardAmount; set => overriddenGiftCardAmount = value; }
        public string RequireOtherProducts { get => requireOtherProducts; set => requireOtherProducts = value; }
        public string RequiredProductIds { get => requiredProductIds; set => requiredProductIds = value; }
        public string AutomaticallyAddRequiredProducts { get => automaticallyAddRequiredProducts; set => automaticallyAddRequiredProducts = value; }
        public string IsDownload { get => isDownload; set => isDownload = value; }
        public string DownloadId { get => downloadId; set => downloadId = value; }
        public string UnlimitedDownloads { get => unlimitedDownloads; set => unlimitedDownloads = value; }
        public string MaxNumberOfDownloads { get => maxNumberOfDownloads; set => maxNumberOfDownloads = value; }
        public string DownloadActivationType { get => downloadActivationType; set => downloadActivationType = value; }
        public string HasSampleDownload { get => hasSampleDownload; set => hasSampleDownload = value; }
        public string SampleDownloadId { get => sampleDownloadId; set => sampleDownloadId = value; }
        public string HasUserAgreement { get => hasUserAgreement; set => hasUserAgreement = value; }
        public string UserAgreementText { get => userAgreementText; set => userAgreementText = value; }
        public string IsRecurring { get => isRecurring; set => isRecurring = value; }
        public string RecurringCycleLength { get => recurringCycleLength; set => recurringCycleLength = value; }
        public string RecurringCyclePeriod { get => recurringCyclePeriod; set => recurringCyclePeriod = value; }
        public string RecurringTotalCycles { get => recurringTotalCycles; set => recurringTotalCycles = value; }
        public string IsRental { get => isRental; set => isRental = value; }
        public string RentalPriceLength { get => rentalPriceLength; set => rentalPriceLength = value; }
        public string RentalPricePeriod { get => rentalPricePeriod; set => rentalPricePeriod = value; }
        public string IsShipEnabled { get => isShipEnabled; set => isShipEnabled = value; }
        public string ShipSeparately { get => shipSeparately; set => shipSeparately = value; }
        public string AdditionalShippingCharge { get => additionalShippingCharge; set => additionalShippingCharge = value; }
        public string DeliveryDate { get => deliveryDate; set => deliveryDate = value; }
        public string IsTaxExempt { get => isTaxExempt; set => isTaxExempt = value; }
        public string TaxCategory { get => taxCategory; set => taxCategory = value; }
        public string IsTelecommunicationsOrBroadcastingOrElectronicServices { get => isTelecommunicationsOrBroadcastingOrElectronicServices; set => isTelecommunicationsOrBroadcastingOrElectronicServices = value; }
        public string ManageInventoryMethod { get => manageInventoryMethod; set => manageInventoryMethod = value; }
        public string UseMultipleWarehouses { get => useMultipleWarehouses; set => useMultipleWarehouses = value; }
        public string WarehouseId { get => warehouseId; set => warehouseId = value; }
        public string StockQuantity { get => stockQuantity; set => stockQuantity = value; }
        public string DisplayStockAvailability { get => displayStockAvailability; set => displayStockAvailability = value; }
        public string DisplayStockQuantity { get => displayStockQuantity; set => displayStockQuantity = value; }
        public string MinStockQuantity { get => minStockQuantity; set => minStockQuantity = value; }
        public string LowStockActivity { get => lowStockActivity; set => lowStockActivity = value; }
        public string NotifyAdminForQuantityBelow { get => notifyAdminForQuantityBelow; set => notifyAdminForQuantityBelow = value; }
        public string BackorderMode { get => backorderMode; set => backorderMode = value; }
        public string AllowBackInStockSubscriptions { get => allowBackInStockSubscriptions; set => allowBackInStockSubscriptions = value; }
        public string OrderMinimumQuantity { get => orderMinimumQuantity; set => orderMinimumQuantity = value; }
        public string OrderMaximumQuantity { get => orderMaximumQuantity; set => orderMaximumQuantity = value; }
        public string AllowedQuantities { get => allowedQuantities; set => allowedQuantities = value; }
        public string AllowAddingOnlyExistingAttributeCombinations { get => allowAddingOnlyExistingAttributeCombinations; set => allowAddingOnlyExistingAttributeCombinations = value; }
        public string NotReturnable { get => notReturnable; set => notReturnable = value; }
        public string DisableBuyButton { get => disableBuyButton; set => disableBuyButton = value; }
        public string DisableWishlistButton { get => disableWishlistButton; set => disableWishlistButton = value; }
        public string AvailableForPreOrder { get => availableForPreOrder; set => availableForPreOrder = value; }
        public string PreOrderAvailabilityStartDateTimeUtc { get => preOrderAvailabilityStartDateTimeUtc; set => preOrderAvailabilityStartDateTimeUtc = value; }
        public string CallForPrice { get => callForPrice; set => callForPrice = value; }
        public string Price { get => price; set => price = value; }
        public string OldPrice { get => oldPrice; set => oldPrice = value; }
        public string ProductCost { get => productCost; set => productCost = value; }
        public string SpecialPrice { get => specialPrice; set => specialPrice = value; }
        public string SpecialPriceStartDateTimeUtc { get => specialPriceStartDateTimeUtc; set => specialPriceStartDateTimeUtc = value; }
        public string SpecialPriceEndDateTimeUtc { get => specialPriceEndDateTimeUtc; set => specialPriceEndDateTimeUtc = value; }
        public string CustomerEntersPrice { get => customerEntersPrice; set => customerEntersPrice = value; }
        public string MinimumCustomerEnteredPrice { get => minimumCustomerEnteredPrice; set => minimumCustomerEnteredPrice = value; }
        public string MaximumCustomerEnteredPrice { get => maximumCustomerEnteredPrice; set => maximumCustomerEnteredPrice = value; }
        public string BasepriceEnabled { get => basepriceEnabled; set => basepriceEnabled = value; }
        public string BasepriceAmount { get => basepriceAmount; set => basepriceAmount = value; }
        public string BasepriceUnit { get => basepriceUnit; set => basepriceUnit = value; }
        public string BasepriceBaseAmount { get => basepriceBaseAmount; set => basepriceBaseAmount = value; }
        public string BasepriceBaseUnit { get => basepriceBaseUnit; set => basepriceBaseUnit = value; }
        public string MarkAsNew { get => markAsNew; set => markAsNew = value; }
        public string MarkAsNewStartDateTimeUtc { get => markAsNewStartDateTimeUtc; set => markAsNewStartDateTimeUtc = value; }
        public string MarkAsNewEndDateTimeUtc { get => markAsNewEndDateTimeUtc; set => markAsNewEndDateTimeUtc = value; }
        public string Weight { get => weight; set => weight = value; }
        public string Length { get => length; set => length = value; }
        public string Width { get => width; set => width = value; }
        public string Height { get => height; set => height = value; }
        public string Categories { get => categories; set => categories = value; }
        public string Manufacturers { get => manufacturers; set => manufacturers = value; }
        public string Picture1 { get => picture1; set => picture1 = value; }
        public string Picture2 { get => picture2; set => picture2 = value; }
        public string Picture3 { get => picture3; set => picture3 = value; }
        public string IsFreeShipping { get => isFreeShipping; set => isFreeShipping = value; }


        public string Style { get => style; set => style = value; }
        public string Country { get => country; set => country = value; }
        public string Ingredients { get => ingredients; set => ingredients = value; }
        public string Attribute { get => attribute; set => attribute = value; }
        public string Bottlesize { get => bottlesize; set => bottlesize = value; }
        public string Containertype { get => containertype; set => containertype = value; }
        public string Vintage { get => vintage; set => vintage = value; }
        public string Region { get => region; set => region = value; }

        //set default product value
        public void Initializer()
        {
            ProductType = "Simple Product";
            VisibleIndividually = "True";
            Vendor = "Bibendum";
            ProductTemplate = "Simple product";
            ShowOnHomePage = "False";
            AllowCustomerReviews = "True";
            Published = "True";
            IsGiftCard = "False";
            GiftCardType = "Virtual";
            OverriddenGiftCardAmount = "0";
            RequireOtherProducts = "False";
            AutomaticallyAddRequiredProducts = "False";
            IsDownload = "False";
            DownloadId = "0";
            UnlimitedDownloads = "True";
            MaximumCustomerEnteredPrice = "10";
            DownloadActivationType = " When Order Is Paid";
            HasSampleDownload = "False";
            SampleDownloadId = "0";
            HasUserAgreement = "False";
            IsRecurring = "False";
            RecurringCycleLength = "100";
            RecurringCyclePeriod = "Days";
            IsShipEnabled = "True";
            IsFreeShipping = "False";
            ShipSeparately = "False";
            AdditionalShippingCharge = "0";
            IsTaxExempt = "False";
            TaxCategory = "GST";
            IsTelecommunicationsOrBroadcastingOrElectronicServices = "False";
            ManageInventoryMethod = "Manage Stock";
            UseMultipleWarehouses = "False";
            WarehouseId = "0";
            DisplayStockQuantity = "True";
            LowStockActivity = "Nothing";
            BackorderMode = " No Backorders";
            AllowBackInStockSubscriptions = "False";
            OrderMaximumQuantity = "10000";
            AllowAddingOnlyExistingAttributeCombinations = "False";
            NotReturnable = "False";
            DisableWishlistButton = "False";
            AvailableForPreOrder = "False";
            CallForPrice = "False";
            OldPrice = "0";
            ProductCost = "0";
            CustomerEntersPrice = "False";
            MinimumCustomerEnteredPrice = "0";
            MaximumCustomerEnteredPrice = "1000";
            BasepriceEnabled = "False";
            BasepriceBaseAmount = "0";
            BasepriceBaseUnit = "bottle(s)";
            MarkAsNew = "False";
            Weight = "1";
            Length = "0";
            Width = "0";
            Height = "0";

            Categories = "Wine";
            Style = "Red Wine";
            Country = "Australia";
            Region = "Barossa Valley";
            Bottlesize = "500ml";
            Containertype = "bottle";
            Vintage = "2017";



        }
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
