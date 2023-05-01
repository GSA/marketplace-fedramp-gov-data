/**
 * 
 * Maintain this script with github.com/gsa/marketplace-fedramp-gov-data
 * 
 * In "Project Settings" (gear icon to left), maintain the script properties variable for the account publishing updates.
 * 
 */

/**
 * In order to protect itself from humans adding/updating/reordering/removing sheets and their columns
 * this script validates each sheet and their columns before processing.  The errSheet and errArray
 * help eliminate repetitive error messages.
 */
var errSheet = "";
var errArr = [];

/**
 * Using a regular expression to scrub data into html class names (used to filter
 * list data on the /products, /agencies, and /assessors pages).  Some strings contain
 * characters that are not allowed in class names.
 * 
 * /[\W_]+/g:
 *    /  = regex change delimiter.
 *    [] = Set of characters to match.
 *    \W = Any "non-word" character, opposite of [^a-zA-z0-9_]
 *    _  = Also match on _ to keep consistent.
 *    +  = Match previous character(s) one-many times.
 *    /g = Global. Do for entire string (don't return after first occurance).
 */
const REGEX = /[\W_]+/g;

/**
 * Sheet names go here here.  Add new sheet names and their columns to isValidSetup() function.
 */
const MASTER_AUTHORIZATION_STATUS_SHEET = "Master Authorization Status";
const MASTER_AGENCY_TAB_SHEET           = "Master Agency Tab";
const MASTER_3PAO_LIST_SHEET            = "Master 3PAO List";
const METRICS_SHEET                     = "Metrics";

/**
 * Arrays of sheet columns for each sheet above
 */
const MASTER_AUTHORIZATION_STATUS_HEADERS = ["FR ID#",	"CSP",	"CSO",	"Service Model",	"Authorization Type",	"Deployment Model",	"Impact Level",	"UEI Number",	"Current 3PAO ID",	"Security Contact Email",	"Sales Contact Email",	"Small Business?",	"Logo URL",	"CSP Website",	"CSO Description",	"CSP Business Function",	"In Process Initial Authorization Agency ID",	"In Process Initial Sub Agency ID",	"Current Active Status?",	"FR Ready Active?",	"FRR Most Recent Date",	"In Process JAB Review Active?",	"In Process JAB Review Most Recent Date",	"In Process Agency Review Active?",	"In Process Agency Review Most Recent Date",	"FedRAMP In Process PMO Review Active?",	"FedRAMP In Process PMO Review Most Recent Date",	"FedRAMP Authorized Active?",	"Non-Recent Authorized Services",	"Recently Updated Authorized Services",	"Authorizations",	"Reuse",	"Agency Authorizations",	"Reuse Agencies",	"Leveraged Systems", "Annual Assessment"];
const MASTER_AGENCY_TAB_HEADERS = ["Agency ID",	"Agency Name",	"Sub Agency",	"E-mail",	"Logo URL",	"Website",	"Authorizations",	"Authorizations Number",	"Reuse",	"Reuse Number",	"In Process (Agency Review)",	"In Process (FedRAMP Review)","In Process (JAB Review)"];
const MASTER_3PAO_LIST_HEADERS = ["3PAO ID#",	"Cert #",	"3PAO Name",	"POC Name",	"POC Email",	"Date Applied",	"A2LA Accreditation Date",	"FedRAMP Accreditation Date",	"Logo URL",	"Year Company Founded",	"Website URL",	"Primary Office Locations",	"Description of 3PAO Services",	"Consulting Services?",	"Description of Consulting Services",	"Additional Cyber Frameworks Your Company Is Accredited to Perform",	"Active?",	"CSPs providing consulting service to",	"Current Clients",	"Products Assessing (Number)"];
const METRICS_HEADERS = ["FR ID", 	"Reuse ATOs", 	"Total ATOs", 	"Indirect Reuse", 	"ATOs", 	"Direct Reuse", 	"Total ATOs", 	"Indirect Reuse", 	"Ready", 	"In Process", 	"Authorizations"];


/**
 * Constants helping the "Metrics" "latest" list.  These constants are used to help eliminate unset dates.
 */
const NO_FRR = "No FRR Date";
const NO_JAB = "Not Active";
const NO_AGENCY = "Not Agency Partnered";
const NO_PMO = "Not In Process";
const NO_AUTH = "Not Authorized";

const STATUS_FRR = "Now FedRAMP Ready";
const STATUS_JAB = "Now in JAB Review";
const STATUS_AGENCY = "Now in Agency Review";
const STATUS_PMO = "Now in PMO Review";
const STATUS_AUTH = "Now Authorized";

                      // ----+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
const STRING_80_BYTES = "                                                                                ";

/***************************************************************************************************/


/**
 * GitHub personal access token (https://github.com/blog/1509-personal-api-tokens)
 */
const github = {
  'owner': 'GSA',
  'repo': 'marketplace-fedramp-gov-data',
  'path': 'data.json',
  'branch': 'master',
  'accessToken': PropertiesService.getScriptProperties().getProperty('GIT_ACCESS_TOKEN'),
  'commitMessage': Utilities.formatString('Published on %s', Utilities.formatDate(new Date(), 'America/New_York', 'yyyy-MM-dd HH:mm:ss'))
};

/***************************************************************************************************/


/**
 * 
 */
function main() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if(isValidSetup(ss) == false) {
    return;
  }

  //createJson(ss);
  // l(createJson(ss));
  updateGitHubRepo(getGitHubSha(), createJson(ss));     // Create json, retrieve sha, update github
}

/***************************************************************************************************/


/**
 * 
 * 
 * 
 */
function createJson(ss) {

  l("Building JSON...");

  var json = {
    meta: { 
      last_change: Utilities.formatDate(new Date(), 'America/New_York', 'yyyy-MM-dd\'T\'HH:mm:ss.SSS\'Z\''),
      produced_by: "General Services Administration"
    },
    data: {
      "Metrics": {
        total: "", 
        latest: []
      },
      "Filters": {
        "product": {},
        "agency": {},
        "assessor":{}
      },
      "Products": [],
      "Agencies": [],
      "Assessors": []
      }
  }


  /**
   * Metrics 
   * 
   * Note: "latest" metric built as part of "products"
   * 
   */

  l("   Metrics");
  initErrorCols(METRICS_SHEET);

  var metricsVals = ss.getSheetByName(METRICS_SHEET).getDataRange().getValues();
  json.data.Metrics.total = metricsVals[1][getCol(METRICS_HEADERS, "Authorizations")];

  /***************************************************************************************************/


  /**
   * Master Authorization Status
   */

  var product = {

    filter_classes: String,

    id: String,
    name: String,
    csp: String,
    cso: String,
    logo: String,
    service_offering: String,
    status: String,
    authorization: String,
    reuse: String,

    // TODO - Maybe?
    //a_status: String,
    //f_process: String,

    a_ready_status: String,
    a_ready_date: String,
    a_ip_jab_status: String,
    a_ip_jab_date: String,
    a_ip_agency_status: String,
    a_ip_agency_date: String,
    a_ip_pmo_status: String,
    a_ip_pmo_date: String,
    a_auth_date: String,

    auth_type: String,
    fedramp_ready: String,

    partnering_agency: String,

    fedramp_auth: String,
    annual_assessment: String,
    independent_assessor: String,
    service_model: [],
    deployment_model: String,
    impact_level: String,
    
    leveraged_systems: String,

    agency_authorizations: [],
    agency_reuse: [],

    service_desc: String,
    sales_email: String,
    security_email: String,
    website: String,
    uei: String,
    small_business: String,
    business_function: [],
    
    service_last_90: [],
    all_others: []
  }
  
  l("   Products");
  initErrorCols(MASTER_AUTHORIZATION_STATUS_SHEET);

  var workLatest;
  var workArr = [];

  var masterAuthVals = ss.getSheetByName(MASTER_AUTHORIZATION_STATUS_SHEET).getDataRange().getValues();

  /**
   * Loop through values on this tab, filling fields.  Special processing generally happens at the bottom of the loop.
   */
  for(var i = 1; i < masterAuthVals.length && 
      masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "FR ID#")] != "" &&
      masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "CSP")] != ""; i++) {
  
    product = {};

    product.id = masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "FR ID#")];
    product.name = masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "CSP")];
    product.csp = product.name; // To utilize sort
    product.cso = masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "CSO")];

    product.logo = masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "Logo URL")];
    product.service_offering = masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "CSO")];
    product.status = masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "Current Active Status?")];
    product.authorization = masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "Authorizations")];
    product.reuse = masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "Reuse")];

    product.ready_status = masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "FR Ready Active?")];
    product.ready_date = masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "FRR Most Recent Date")];
    product.ip_jab_status = masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "In Process JAB Review Active?")];
    product.ip_jab_date = masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "In Process JAB Review Most Recent Date")];
    product.ip_agency_status = masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "In Process Agency Review Active?")];
    product.ip_agency_date = masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "In Process Agency Review Most Recent Date")];
    product.ip_pmo_status = masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "FedRAMP In Process PMO Review Active?")];
    product.ip_pmo_date = masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "FedRAMP In Process PMO Review Most Recent Date")];
    product.auth_date = masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "FedRAMP Authorized Active?")];

    product.auth_type = masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "Authorization Type")];
    product.fedramp_ready = masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "FRR Most Recent Date")];

    product.partnering_agency =  masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "In Process Initial Authorization Agency ID")];

    if(masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "In Process Initial Sub Agency ID")] != "" ) {
      product.partnering_agency =  masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "In Process Initial Sub Agency ID")];
    }

    product.fedramp_auth = masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "FedRAMP Authorized Active?")];
    product.annual_assessment = masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "Annual Assessment")];
    product.independent_assessor = masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "Current 3PAO ID")];
    product.service_model = Array.from(new Set(masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "Service Model")].split("|"))).sort();
    product.deployment_model = masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "Deployment Model")];
    product.impact_level = masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "Impact Level")];
    product.impact_level_number = getImpactLevelNumber(product.impact_level);

    // Build and sort before storing
    workArr = getItems(masterAuthVals, Array.from(new Set(masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "Leveraged Systems")].split("|"))));
    product.leveraged_systems = quickSortOnObjectCSP(workArr, 0, workArr.length-1);

    // Force array into Set() to dedup, then sort before storing
    product.agency_authorizations = Array.from(new Set(masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "Agency Authorizations")].split("|"))).sort();
    
    // Force array into Set() to dedup, then sort before storing
    product.agency_reuse = Array.from(new Set(masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "Reuse Agencies")].split("|"))).sort();

    product.service_desc = masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "CSO Description")];
    product.sales_email = masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "Sales Contact Email")];
    product.security_email = masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "Security Contact Email")];
    product.website = masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "CSP Website")];
    product.uei = masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "UEI Number")];
    product.small_business = getYesNo(masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "Small Business?")]);
    product.business_function = Array.from(new Set(masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "CSP Business Function")].split("|"))).sort();

    product.service_last_90 = Array.from(new Set(masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "Recently Updated Authorized Services")].split("|"))).sort();
    product.all_others = Array.from(new Set(masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "Non-Recent Authorized Services")].split("|"))).sort();

    /**
     * Special processing for finding the 3 "latest" dates from all the below columns.
     */
    workLatest = getLatest(product.id, product.logo, product.service_offering, 
          masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "FRR Most Recent Date")],
          masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "In Process JAB Review Most Recent Date")],
          masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "In Process Agency Review Most Recent Date")],
          masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "FedRAMP In Process PMO Review Most Recent Date")],
          masterAuthVals[i][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "FedRAMP Authorized Active?")]);

    if(workLatest.date != "") {                   // If there's a good date in any of the above columns...
    
      json.data.Metrics.latest.push(workLatest);  // Build array of "latest" objects to be paired down later.
    }
    
    // Filter classes
    product.filter_classes = "";
    product.filter_classes += " filter-status-" + product.status.replace(REGEX,"-");
    product.filter_classes += " filter-impact-level-" + product.impact_level.replace(REGEX,"-");
    product.filter_classes += " filter-auth-type-" + product.auth_type.replace(REGEX,"-");
    product.filter_classes += " filter-deployment-model-" + product.deployment_model.replace(REGEX,"-");
    product.filter_classes += " filter-small-business-" + product.small_business.replace(REGEX,"-");
    product.filter_classes += " filter-assessor-" + product.independent_assessor.replace(REGEX,"-");

    for(var j = 0; j < product.service_model.length; j++) {
      product.filter_classes += " filter-service-model-" + product.service_model[j].replace(REGEX,"-");
    }

    for(var j = 0; j < product.business_function.length; j++) {
      product.filter_classes += " filter-business-function-" + product.business_function[j].replace(REGEX,"-");
    }

    json.data.Products.push(product);             // Build array of Products
  }

  json.data.Products = quickSortOnObjectCSP(json.data.Products, 0, json.data.Products.length-1); // Sort

  /**
   * This object should ALWAYS contain more than 3 items before the splice(). If this
   * line blows up, there are bigger problems.
   */
  json.data.Metrics.latest = quickSortOnObjectDate(json.data.Metrics.latest, 0, json.data.Metrics.latest.length-1).splice(0,3);
  
  /***************************************************************************************************/


  /**
   * Agencies
   */

  var agency = {

    filter_classes: String,

    id: String,
    parent: String,
    sub: String,
    csp: String,
    logo: String,
    authorization: Number,
    reuse: Number,
    email: String,
    website: String,
    auths: [],
    reuses: [],
    procs: []
  }

  l("   Agencies");
  initErrorCols(MASTER_AGENCY_TAB_SHEET);

  var masterAgencyVals = ss.getSheetByName(MASTER_AGENCY_TAB_SHEET).getDataRange().getValues();

  /**
   * Loop through values on this tab, filling fields.  Special processing generally happens at the bottom of the loop.
   */
  for(var i = 1; i < masterAgencyVals.length && masterAgencyVals[i][getCol(MASTER_AGENCY_TAB_HEADERS,"Agency ID")] != ""; i++) {
  
    agency = {};

    agency.id = masterAgencyVals[i][getCol(MASTER_AGENCY_TAB_HEADERS,"Agency ID")];
    agency.parent = masterAgencyVals[i][getCol(MASTER_AGENCY_TAB_HEADERS, "Agency Name")];
    agency.sub = masterAgencyVals[i][getCol(MASTER_AGENCY_TAB_HEADERS, "Sub Agency")];
    
    if(agency.parent == agency.sub) {  // Sub and Parent the same?

      agency.sub = "";                 // Remove. We don't want it displayed twice on the webpage.
    }

    agency.csp = concatParentSub(agency.parent, agency.sub);

    agency.logo = masterAgencyVals[i][getCol(MASTER_AGENCY_TAB_HEADERS, "Logo URL")];
    agency.authorization = masterAgencyVals[i][getCol(MASTER_AGENCY_TAB_HEADERS, "Authorizations Number")];
    agency.reuse = masterAgencyVals[i][getCol(MASTER_AGENCY_TAB_HEADERS, "Reuse Number")];
    agency.email = masterAgencyVals[i][getCol(MASTER_AGENCY_TAB_HEADERS, "E-mail")];
    agency.website = masterAgencyVals[i][getCol(MASTER_AGENCY_TAB_HEADERS, "Website")];

    // Build and sort before storing
    workArr = getItems(masterAuthVals, Array.from(new Set(masterAgencyVals[i][getCol(MASTER_AGENCY_TAB_HEADERS, "Authorizations")].split("|"))));
    agency.auths = quickSortOnObjectCSP(workArr, 0, workArr.length-1);

    // Build and sort before storing
    workArr = getItems(masterAuthVals, Array.from(new Set(masterAgencyVals[i][getCol(MASTER_AGENCY_TAB_HEADERS, "Reuse")].split("|"))));
    agency.reuses = quickSortOnObjectCSP(workArr, 0, workArr.length-1);

    // Build and sort before storing.  There's a concat snuck in here, too.
    workArr = getItems(masterAuthVals, Array.from(new Set(masterAgencyVals[i][getCol(MASTER_AGENCY_TAB_HEADERS, "In Process (Agency Review)")].split("|"))), "Agency Review").concat(getItems(masterAuthVals, Array.from(new Set(masterAgencyVals[i][getCol(MASTER_AGENCY_TAB_HEADERS, "In Process (FedRAMP Review)")].split("|"))),"FedRAMP Review"), getItems(masterAuthVals, Array.from(new Set(masterAgencyVals[i][getCol(MASTER_AGENCY_TAB_HEADERS, "In Process (JAB Review)")].split("|"))),"JAB Review"));
    agency.procs = quickSortOnObjectCSP(workArr, 0, workArr.length-1);


    // Filter classes
    agency.filter_classes = "";
    agency.filter_classes += " filter-parent-agency-" + agency.parent.replace(REGEX,"-");
    agency.filter_classes += getFilterClassBucket("authorization", agency.authorization);
    agency.filter_classes += getFilterClassBucket("reuse", agency.reuse);
    agency.filter_classes += getFilterClassImpactAndOffering(Array.from(new Set(agency.auths.concat(agency.procs))));

    json.data.Agencies.push(agency);    // Build array of Agencies
  }

  json.data.Agencies = quickSortOnObjectCSP(json.data.Agencies, 0, json.data.Agencies.length-1); 

  /***************************************************************************************************/


  /**
   * Assessors
   */

  var assessor = {

    filter_classes: String,

    id: String,
    name: String,
    csp: String,
    logo: String,
    products_assessing: Number,
    accredited_since: String,
    poc: String,
    email: String,
    founded: String,
    address: String,
    desc: String,
    services: String,
    csps: String,
    frameworks: String,
    clients: []
  }

  l("   Assessors");
  initErrorCols(MASTER_3PAO_LIST_SHEET);

  var master3paoVals = ss.getSheetByName(MASTER_3PAO_LIST_SHEET).getDataRange().getValues();

  /**
   * Loop through values on this tab, filling fields.  Special processing generally happens at the bottom of the loop.
   */
  for(var i = 1; i < master3paoVals.length && master3paoVals[i][getCol(MASTER_3PAO_LIST_HEADERS, "3PAO ID#")] != ""; i++) {
  
    assessor = {};

    assessor.id = master3paoVals[i][getCol(MASTER_3PAO_LIST_HEADERS, "3PAO ID#")];
    assessor.name = master3paoVals[i][getCol(MASTER_3PAO_LIST_HEADERS, "3PAO Name")];
    assessor.csp = assessor.name;

    assessor.logo = master3paoVals[i][getCol(MASTER_3PAO_LIST_HEADERS, "Logo URL")];
    assessor.products_assessing = master3paoVals[i][getCol(MASTER_3PAO_LIST_HEADERS, "Products Assessing (Number)")];

    assessor.accredited_since = master3paoVals[i][getCol(MASTER_3PAO_LIST_HEADERS, "A2LA Accreditation Date")];
    assessor.poc = master3paoVals[i][getCol(MASTER_3PAO_LIST_HEADERS, "POC Name")];
    assessor.email = master3paoVals[i][getCol(MASTER_3PAO_LIST_HEADERS, "POC Email")];
    assessor.founded = master3paoVals[i][getCol(MASTER_3PAO_LIST_HEADERS, "Year Company Founded")];
    assessor.address = master3paoVals[i][getCol(MASTER_3PAO_LIST_HEADERS, "Primary Office Locations")];
    assessor.desc = master3paoVals[i][getCol(MASTER_3PAO_LIST_HEADERS, "Description of 3PAO Services")];
    
    if(master3paoVals[i][getCol(MASTER_3PAO_LIST_HEADERS, "Consulting Services?")] == 'Y') {

      assessor.services = master3paoVals[i][getCol(MASTER_3PAO_LIST_HEADERS, "Description of Consulting Services")];
    }

    assessor.csps = master3paoVals[i][getCol(MASTER_3PAO_LIST_HEADERS, "CSPs providing consulting service to")];
    assessor.frameworks = master3paoVals[i][getCol(MASTER_3PAO_LIST_HEADERS, "Additional Cyber Frameworks Your Company Is Accredited to Perform")];
    
    // Build and sort before storing
    workArr = getItems(masterAuthVals, Array.from(new Set(master3paoVals[i][getCol(MASTER_3PAO_LIST_HEADERS, "Current Clients")].split("|"))));
    assessor.clients = quickSortOnObjectCSP(workArr, 0, workArr.length-1);

    // Filter classes
    assessor.filter_classes = "";
    assessor.filter_classes += getFilterClassBucket("products-assessing", assessor.products_assessing);
    assessor.filter_classes += getFilterClassImpactAndOffering(assessor.clients);

    json.data.Assessors.push(assessor);    // Build array of Assessors
  }

  json.data.Assessors = quickSortOnObjectCSP(json.data.Assessors, 0, json.data.Assessors.length-1); 

  /***************************************************************************************************/

  /**
   * Product Filters
   */

  l("   Product  Filters");

  var filter = {
    name: String,
    class_name: String 
  }
  var productFilters = {
    status: [],
    business_function: [],
    service_model: [],
    impact_level: [],
    auth_type: [],
    deployment_models: [],
    small_business: [],
    assessor: []
  }

  productFilters.status = getFilter("status", masterAuthVals, MASTER_AUTHORIZATION_STATUS_HEADERS, "Current Active Status?");
  productFilters.business_function = getFilter("business-function", masterAuthVals, MASTER_AUTHORIZATION_STATUS_HEADERS, "CSP Business Function");
  productFilters.service_model = getFilter("service-model", masterAuthVals, MASTER_AUTHORIZATION_STATUS_HEADERS, "Service Model"); 

  productFilters.impact_level.push({ name: "LI-SaaS", class_name: "filter-impact-level-LI-SaaS"});
  productFilters.impact_level.push({ name: "Low",     class_name: "filter-impact-level-Low"});
  productFilters.impact_level.push({ name: "Moderate",class_name: "filter-impact-level-Moderate"});
  productFilters.impact_level.push({ name: "High",    class_name: "filter-impact-level-High"});

  productFilters.auth_type = getFilter("auth-type", masterAuthVals, MASTER_AUTHORIZATION_STATUS_HEADERS, "Authorization Type");
  productFilters.deployment_models = getFilter("deployment-model", masterAuthVals, MASTER_AUTHORIZATION_STATUS_HEADERS, "Deployment Model"); 

  // productFilters.small_business.push({ name: "Yes", class_name: "filter-small-business-Yes"});
  // productFilters.small_business.push({ name: "No", class_name: "filter-small-business-No"});

  productFilters.assessor = getFilter("assessor", masterAuthVals, MASTER_AUTHORIZATION_STATUS_HEADERS, "Current 3PAO ID");

  json.data.Filters.product = productFilters;

  /***************************************************************************************************/


  /**
   * Agency Filters
   */

  l("   Agency   Filters");

  var agencyFilters = {
    parent_agency: [],
    authorization: [], 
    reuse: [],
    impact_level: [] 
  }  

  agencyFilters.parent_agency = getFilter("parent-agency", masterAgencyVals, MASTER_AGENCY_TAB_HEADERS, "Agency Name"); 

  // if these numbers change, also change logic in getFilterClassBucket
  agencyFilters.authorization.push({ name: "0-5", class_name: "filter-authorization-1"});
  agencyFilters.authorization.push({ name: "6-10", class_name: "filter-authorization-2"});
  agencyFilters.authorization.push({ name: "11-20", class_name: "filter-authorization-3"});
  agencyFilters.authorization.push({ name: "21+", class_name: "filter-authorization-4"});


  agencyFilters.reuse.push({ name: "0-5", class_name: "filter-reuse-1"});
  agencyFilters.reuse.push({ name: "6-10", class_name: "filter-reuse-2"});
  agencyFilters.reuse.push({ name: "11-20", class_name: "filter-reuse-3"});
  agencyFilters.reuse.push({ name: "21+", class_name: "filter-reuse-4"});


  agencyFilters.impact_level.push({ name: "LI-SaaS", class_name: "filter-impact-level-LI-SaaS"});
  agencyFilters.impact_level.push({ name: "Low",     class_name: "filter-impact-level-Low"});
  agencyFilters.impact_level.push({ name: "Moderate",class_name: "filter-impact-level-Moderate"});
  agencyFilters.impact_level.push({ name: "High",    class_name: "filter-impact-level-High"});

  json.data.Filters.agency = agencyFilters;

  /***************************************************************************************************/


  /**
   * Assessor Filters
   */

  l("   Assessor Filters");

  var assessorFilters = {
    product_assessing: [],
    impact_level: [],
    status: []
  }

  assessorFilters.product_assessing.push({ name: "0-5", class_name: "filter-products-assessing-1"});
  assessorFilters.product_assessing.push({ name: "6-10", class_name: "filter-products-assessing-2"});
  assessorFilters.product_assessing.push({ name: "11-20", class_name: "filter-products-assessing-3"});
  assessorFilters.product_assessing.push({ name: "21+", class_name: "filter-products-assessing-4"});

  assessorFilters.impact_level = json.data.Filters.product.impact_level;
  assessorFilters.status = json.data.Filters.product.status;

  json.data.Filters.assessor = assessorFilters;

  /***************************************************************************************************/

  return JSON.stringify(json);
}

/***************************************************************************************************/


/**
 * Find the latest date from the last 5 fields passed in (if any of them have dates).
 * 
 * @param {frr} - Possible date (FedRAMP Ready)
 * @param {jab} - Possible date (JAB Review)
 * @param {agency} - Possible date (Agency Review)
 * @param {pmo} - Possible date (PMO Review)
 * @param {auth} - Possible date (Auth Review)
 * @returns {Object} - Latest date and its "status" (description of date)
 */
function getLatest(id, logo, cso, frr, jab, agency, pmo, auth) {

  var rec = {
    id: String,
    logo: String,
    cso: String,
    date: String,
    status: String
  }

  // No good dates? Return.
  if(frr == NO_FRR && jab == NO_JAB && agency == NO_AGENCY && pmo == NO_PMO && auth == NO_AUTH) {

    return rec;
  }

  rec.id = id;
  rec.logo = logo;
  rec.cso = cso;

  // Find latest date
  var item = maxDate(maxDate(maxDate(maxDate(
    {date: scrubNo(frr),    status: STATUS_FRR}, 
    {date: scrubNo(jab),    status: STATUS_JAB}), 
    {date: scrubNo(agency), status: STATUS_AGENCY}), 
    {date: scrubNo(pmo),    status: STATUS_PMO}), 
    {date: scrubNo(auth),   status: STATUS_AUTH});

  rec.date = item.date;
  rec.status = item.status;

  return rec;
}

/***************************************************************************************************/


/**
 * This utility function is used to clear any default literals within a date field.
 * 
 * @param {inDate} - Possible detault literal or actual date.
 * @returns {String} - Empty or the date passed in.
 */
function scrubNo(inDate) {

  if(inDate == NO_FRR || inDate == NO_JAB || inDate == NO_AGENCY || inDate == NO_PMO || inDate == NO_AUTH) {

    return "";
  }
  return inDate;
}

/***************************************************************************************************/


/**
 * Returns the greater date between two Objects containing a .date variable.
 * 
 * @param {a} - Object of Date and Status (date description)
 * @param {b} - Object of Date and Status (date description)
 * @returns {Object} - Date and Status of "latest" date
 */
function maxDate(a, b) {

  if(a.date > b.date) {

    return a;
  }
  return b;
}

/***************************************************************************************************/


/**
 * Get unique sorted list of all values present in a column.
 * 
 * @param {inLabel} - Label to be applied to the filter so each category is unique.
 * @param {inVals} - Values from a sheet
 * @param {inHeaders} - Array of headers from a sheet
 * @param {inCol} - Column 
 * @returns {Object} - 
 */
function getFilter(inLabel, inVals, inHeaders, inCol) {
  
  var filter = {
    name: String,
    class_name: String 
  }

  var filterArr = [];
  
  for(var i = 1; i < inVals.length; i++) {
  
    filterArr = filterArr.concat(inVals[i][getCol(inHeaders, inCol)].split("|"));   // Some columns are pipe-delimited
  }

  // The easiest way to dedup an array is to force it into a Set
  filterArr = Array.from(new Set(filterArr)).sort();

  var filters = [];

  for(var i = 0; i < filterArr.length; i++) {
    
    filter = {};
    
    filter.name = filterArr[i];
    filter.class_name = "filter-" + inLabel + "-" + filterArr[i].replace(REGEX, "-");

    filters.push(filter);
  }

  return filters;

}

/***************************************************************************************************/


/**
 * This function initializes the error sheet name and array to 
 * the sheet name about to be processed. The array will ultimately hold
 * unique sheet/col info so we don't get hundreds of identical messages
 * about columns missing.
 *
 * @param {inSheet} - active spreadsheet being processed
 */
function initErrorCols(inSheet) {

  errSheet = inSheet;
  errArr = [];
}

/***************************************************************************************************/


/**
 * Prints a unique error message and adds it to an array so
 * it won't be duplicated.  This uses the global sheet name
 * set in initErrorCols() above.
 *
 * @param {inCol} - column not found
 */
 function addErrorCol(inCol) {

  var errMsg = "Sheet<" + errSheet + "> Col<" + inCol + ">";

  if(errArr.indexOf(errMsg) == -1) {  // Unique message?

    l(errMsg);                        // Print it
    errArr.push(errMsg);              // Add it to the error array so a future duplicate won't be printed.
  }
}

/***************************************************************************************************/


/**
 * Attempt to find a column within an array of header names.
 * 
 * @param {inHeaders} - Array of header names
 * @param {inCol} - index of column to find
 * @returns {Integer} - index of column found in array of headers. -1 if not found.
 */
function getCol(inHeaders, inCol) {

  var colNum = inHeaders.indexOf(inCol);  // Find index of column in header array

  if(colNum == -1) {                      // Not found?

    addErrorCol(inCol);                   // Print unique missing columns
  }
  return colNum;                          // Return index
}

/***************************************************************************************************/


/**
 * In several locations of the AuthLog cells contain a pipe-delimited list of
 * IDs that are used to create an array of objects containing info from
 * corresponding entries on the Master Authorization Status sheet.
 * 
 * @param {masterAuthVals} - 2D array of values from the Master Authorization Status sheet so they don't have to be retrieved multiple times.
 * @param {itemArray} - Array of Master Authorization Status IDs
 * @returns {Array of Objects} - Array of "item" objects
 */
function getItems(masterAuthVals, itemArray, inConst = "") {

  var item = {
    id: String,
    csp: String, 
    cso: String, 
    status: String,
    impact: String
  }

  var list = [];

  /**
   * Loop through each ID from the array
   */
  for(var j = 0; j < itemArray.length; j++) { 

    /**
     * The itemArray is created from a string being split(). The result will 
     * be at least 1 item (even if it's splitting a null field). Check for blank before
     * looping through everything.
     */
    if(itemArray[j] != "") {

      /**
       * Loop through the Master Authorization Status rows, looking for the ID.
       */
      for(var k = 1; k < masterAuthVals.length; k++) {

        /**
         * If ID matches, push an item onto the array
         */
        if(itemArray[j] == masterAuthVals[k][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "FR ID#")]) {

          item = {}
          item.id = masterAuthVals[k][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "FR ID#")];
          item.csp = masterAuthVals[k][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "CSP")];
          item.cso = masterAuthVals[k][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "CSO")];

          if(inConst != "") {
            item.status = inConst;
          } else {
            item.status = masterAuthVals[k][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "Current Active Status?")];
          }

          item.impact_level = masterAuthVals[k][getCol(MASTER_AUTHORIZATION_STATUS_HEADERS, "Impact Level")];
          item.impact_level_number = getImpactLevelNumber(item.impact_level);
         
          list.push(item);
          
          break;
        }
      }
    }
  }
  return list;
}

/***************************************************************************************************/


/**
 * Individually validates each sheet and their columns
 *
 * @param {ss} - active spreadsheet
 * @returns {Boolean} - false if any sheet or sheet's column is missing.
 */
function isValidSetup(ss) {

  l("Validating Sheets and Columns");

  if(isValidSheetColumns(ss, MASTER_AUTHORIZATION_STATUS_SHEET, MASTER_AUTHORIZATION_STATUS_HEADERS) == false 
  || isValidSheetColumns(ss, MASTER_AGENCY_TAB_SHEET, MASTER_AGENCY_TAB_HEADERS) == false 
  || isValidSheetColumns(ss, MASTER_3PAO_LIST_SHEET, MASTER_3PAO_LIST_HEADERS) == false
  || isValidSheetColumns(ss, METRICS_SHEET, METRICS_HEADERS) == false) { 

    return false;
  }
  return true;
}

/***************************************************************************************************/


/**
 * Validates columns expected in a sheet
 *
 * @param {ss} - active spreadsheet
 * @param {inName} - name of sheet to validate
 * @param {inHeaders} - array of headers expected for sheet
 * @returns {Boolean} - false if sheet or sheet's column is missing.
 */
function isValidSheetColumns(ss,inName,inHeaders) {
  
  if(ss.getSheetByName(inName) == false) {        // Does the sheet exist?

    l("Missing Sheet Name: <"+ inName +">");      // Print something useful
    return false;                                 // Bail
  }

  // Grab headers from the actual sheet to validate
  var range = ss.getSheetByName(inName).getRange(1,1,1,inHeaders.length);
  var values = range.getValues(); 
 
  // Compare actual headers to the headers we expect.
  for(var i = 0; i < inHeaders.length; i++) { 

    if(values[0][i] != inHeaders[i]) {

      l("Invalid Header\n   Sheet<" + inName +">\n   Change<" + inHeaders[i] +"> --> <"+ values[0][i]+">");
      return false;                                // Print something useful and bail.
    }
  }
  return true;
}

/***************************************************************************************************/


/**
 * Lazy way to print log messages... OR GENIUS!?
 *
 * @param {inMsg} - message to possibly print to log
 */
function l(inMsg) {
  Logger.log(inMsg);
}

/***************************************************************************************************/


/**
 * HTTP PUT to commit an element to github.
 *
 * @param {sha} - sha of element being replaced
 * @param {json} - json file contents to commit to 
 */
function updateGitHubRepo(sha, json) {

  l("Updating GitHub");

  var requestUrl = Utilities.formatString(
    'https://api.github.com/repos/%s/%s/contents/%s',
    github.owner,
    github.repo,
    github.path,
    github.branch
  ), 
  response = UrlFetchApp.fetch(requestUrl, {
    'method': 'PUT',
    'headers': {
       'Accept': 'Accept: application/vnd.github+json',
       'Authorization': Utilities.formatString('Bearer %s', github.accessToken)
    },
    'payload': JSON.stringify({
      'message': github.commitMessage,
      'sha': sha,
      'content': Utilities.base64Encode(json)
    })
  })
}

/***************************************************************************************************/


/**
 * Go get the sha of the github element we want to replace.
 * 
 * @param {inMsg} - message to possibly print to log
 * @returns {String} - sha from the github object
 */
function getGitHubSha() {

  var requestUrl = Utilities.formatString(
    'https://api.github.com/repos/%s/%s/contents/%s',
    github.owner,
    github.repo,
    github.path,
    github.branch
  ),
  response = UrlFetchApp.fetch(requestUrl, {
    'method': 'GET',
    'headers': {
       'Accept': 'Accept: application/vnd.github+json',
       'Authorization': Utilities.formatString('Bearer %s', github.accessToken)
    }
  });

  return JSON.parse(response.getContentText()).sha;
}

/***************************************************************************************************/


/**
 * Standard recursive divide/conquor QuickSort to sort on DATE field
 * of Item objects.
 * 
 * @param {arr} - Array of Item objects to divide into partitions and compare
 * @param {low} - Index of arr's lower bound
 * @param {high} - Index of arr's upper bound
 * @returns {arr} - Array to return
 */
function quickSortOnObjectDate(arr, low, high)
{
  if (low < high) {

    var p = partDate(arr, low, high);
    arr = quickSortOnObjectDate(arr, low, p-1);
    arr = quickSortOnObjectDate(arr, p+1, high);
  }
  return arr;
}

/***************************************************************************************************/


/**
 * Standard partition, pivot, and compare for QuickSort using DATE field of Item object.
 * 
 * @param {arr} - Array of Item objects
 * @returns {arr} - Partitioned Array
 */
function partDate(arr, low, high)
{
  var p = arr[high];
  var i = (low - 1);
  var temp;

  for (var j = low; j <= high - 1; j++) {
  
    if (arr[j].date > p.date) {
  
      i++;
  
      temp = arr[i];
      arr[i] = arr[j];
      arr[j] = temp;
    }
  }
  
  temp = arr[i+1];
  arr[i+1] = arr[high];
  arr[high] = temp;

  return (i + 1);
}

/***************************************************************************************************/


/**
 * Standard recursive divide/conquor QuickSort to sort on CSP field
 * of Item objects.
 * 
 * @param {arr} - Array of Item objects to divide into partitions and compare
 * @param {low} - Index of arr's lower bound
 * @param {high} - Index of arr's upper bound
 * @returns {arr} - Array to return
 */
function quickSortOnObjectCSP(arr, low, high)
{
  if (low < high) {

    var p = partCSP(arr, low, high);
    arr = quickSortOnObjectCSP(arr, low, p-1);
    arr = quickSortOnObjectCSP(arr, p+1, high);
  }
  return arr;
}

/***************************************************************************************************/


/**
 * Standard partition, pivot, and compare for QuickSort using CSP field of Item object.
 * 
 * @param {arr} - Array of Item objects
 * @returns {arr} - Partitioned Array
 */
function partCSP(arr, low, high)
{
  var p = arr[high];
  var i = (low - 1);
  var temp;

  for (var j = low; j <= high - 1; j++) {
  
    if (arr[j].csp.toLowerCase() < p.csp.toLowerCase()) {
  
      i++;
  
      temp = arr[i];
      arr[i] = arr[j];
      arr[j] = temp;
    }
  }
  
  temp = arr[i+1];
  arr[i+1] = arr[high];
  arr[high] = temp;

  return (i + 1);
}

/***************************************************************************************************/


/**
 * Returns "Yes" or "No" from 'Y' or not 'Y' input.  For example, the "Small Business" field contains Y/N
 * but the front end website requires Yes/No.
 * 
 * @param {inChar} - Character to interrogate.
 * @returns {literal} - "Yes" or "No" depending on 
 */

function getYesNo(inChar) {

  if(inChar == "Y") {
  
    return "Yes";
  }
  
  return "No";
}

/***************************************************************************************************/


/**
 * The front end list filtering for products/agencies/assessors requires filtering on number
 * ranges of authorizations and reuses.
 * 
 * @param {label} - Label to apply to the middle of the filter name.
 * @returns {s} - Filter name created.
 */
function getFilterClassBucket(label, num) {

  var s = " filter-" + label + "-";

  if(num <= 5) {

    s += "1";
  } else if (num <= 10) {

    s += "2";
  } else if (num <= 20) {

    s += "3";
  } else {

    s += "4";
  }

  return s;
}

/***************************************************************************************************/


/**
 * Concat strings for sorting.  In order to sort nicely, the "parent" and "sub" are placed into 80 byte buffers.
 * 
 * @param {s1} - String to be concatenated and returned
 * @param {s2} - String to be concatenated and returned
 * @returns - String of two 80 byte strings 
 */
function concatParentSub(s1, s2) {

  s1 = s1 + STRING_80_BYTES;
  s1 = s1.slice(0,80);

  s2 = s2 + STRING_80_BYTES;
  s2 = s2.slice(0,80);

  return s1 + s2;

}

/***************************************************************************************************/


/**
 * Get unique list of impact level and status for filtering agencies/assessors.
 * 
 * @param {arr} - Array of Item objects containing product impact level and status.
 * @returns {literal} - String of unique class names.
 */
function getFilterClassImpactAndOffering(arr) {

  var classArr = [];

  if (arr.length == 0) {
    return "";
  }

  for (var i = 0; i < arr.length; i++) {
    classArr.push(" filter-impact-level-" + arr[i].impact_level.replace(REGEX,"-"));
    classArr.push(" filter-status-" + arr[i].status.replace(REGEX,"-"));
  }

  return Array.from(new Set(classArr)).sort().join('');

}

/***************************************************************************************************/


/**
 * Super secret number to hide in front of the impact_level so that it doesn't sort by alpha
 * 
 * @param {inLevel} - Impact Level value
 * @returns {literal} - 1 thorugh 4 for sorting
 */
function getImpactLevelNumber(inLevel) {

  if(inLevel == "LI-SaaS") {
    return "1";
  }
  if(inLevel == "Low") {
    return "2";
  }
  if(inLevel == "Moderate") {
    return "3";
  }
  return "4";
}

/***************************************************************************************************/

