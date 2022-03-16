#!/usr/bin/env ruby

# download new rates from https://ec.europa.eu/taxation_customs/vat-rates_en
SHEET_NAME = 'List of VAT rates applied'

COUNTRY_COLUMN_NAME = 'Member States'
COUNTRY_CODE_COLUMN_NAME = 'Code'
STANDARD_RATE_COLUMN_NAME = 'Standard Rate'
REDUCED_RATE_COLUMN_NAME = 'Reduced Rate'

require 'roo'
require 'json'

filename = File.join(__dir__, 'vat_rates_en.xlsx')
xlsx = Roo::Spreadsheet.open(filename)
rates = []
json_data = {
  version: '1.0',
  description: 'Parsed VAT rates from https://ec.europa.eu/taxation_customs/vat-rates_en (XLSX file)',
  rates: rates
}

sheet = xlsx.sheet(SHEET_NAME)

title_row = nil
1.upto(sheet.last_row) do |line|
  row = sheet.row(line)
  if title_row.nil? &&
     row.index(COUNTRY_COLUMN_NAME) &&
     row.index(COUNTRY_CODE_COLUMN_NAME) &&
     row.index(STANDARD_RATE_COLUMN_NAME) &&
     row.index(REDUCED_RATE_COLUMN_NAME)
    title_row = row
    next
  end

  next unless title_row

  rates << {
    name: row[title_row.index(COUNTRY_COLUMN_NAME)],
    code: row[title_row.index(COUNTRY_CODE_COLUMN_NAME)],
    country_code: row[title_row.index(COUNTRY_CODE_COLUMN_NAME)],
    periods: [{
      effective_from: Time.at(0),
      rates: {
        reduced: row[title_row.index(REDUCED_RATE_COLUMN_NAME)].to_f,
        standard: row[title_row.index(STANDARD_RATE_COLUMN_NAME)].to_f
      }
    }]
  }
end

json_file = File.join(__dir__, '..', 'vat-rates.json')
File.write(json_file, JSON.pretty_generate(json_data))
