package mairo;

public interface CacmXlsxConstants {
  // general constants
  String SHEET_NAME = "CACM_2.2&2.3 flow-based";
  String GENERAL_INFO = "NTC-based capacity allocation and network utilisation";
  String NTC_BASED_CAPACITY_ALLOCATION_AND_NETWORK_UTILISATION = "NTC-based capacity allocation and network utilisation [CACM 2.2&2.3]";
  String NTC_BASED_CAPACITY_ALLOCATION_AND_NETWORK_UTILISATION_RESULT = "NTC-based capacity allocation and network utilisation - Result [CACM 2.2&2.3]";
  String EMPTY_STRING = "";
  String LB_SPACED = " (";
  String RB = ")";

  // general headers constants
  String RESULT_HEADER = "Result";
  String CONSTRAINT_HEADER = "Constraint";
  String PTDF_HEADER = "PTDF";
  String MTU_HEADER = "MTU";
  String MTU_SUMMARY_HEADER = "MTU Summary";
  String TIME_INTERVAL_HEADER = "Time Interval";
  String REGION_HEADER = "Region";
  String DEFAULT_FONT = "Arial";

  //result table headers
  String RESULT_DIRECTION_HEADER = "Out Area \n > \n In Area";
  String RESULT_NAME_HEADER = "Name";
  String RESULT_EIC_HEADER = "EIC";
  String RESULT_TYPE_HEADER = "Type";
  String RESULT_LOCATION_HEADER = "Location";
  String RESULT_MAX_FLOW_HEADER = "Max flow [MW]";
  String RESULT_MAX_EXCHANGE_HEADER = "Max Exchange [MW]";
  String RESULT_MAX_EXCHANGE_AFTER_REMEDY_HEADER = "Max Exchange \n After Remedy [MW]";
  String RESULT_AAC_HEADER = "AAC";
  String RESULT_TRM_HEADER = "TRM";
  String RESULT_TTC_HEADER = "TTC";

  // constraint table headers
  String CONSTRAINT_ID_HEADER = "Constraint ID";
  String CONSTRAINT_NAME_HEADER = "cb/co name";
  String CONSTRAINT_EIC_HEADER = "cb/co eic";
  String CONSTRAINT_TYPE_HEADER = "cb/co type";
  String CONSTRAINT_LOCATION_HEADER = "cb/co location";
  String CONSTRAINT_IN_AREA_HEADER = "cb/co In Area";
  String CONSTRAINT_OUT_AREA_HEADER = "cb/co Out Area";
  String CONSTRAINT_FREF_HEADER = "Fref [MW]";
  String CONSTRAINT_FMAX_HEADER = "Fmax [MW]";
  String CONSTRAINT_ACTUAL_FLOW_HEADER = "Actual flow [MW]";
}
