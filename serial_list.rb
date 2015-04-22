require 'bundler/setup'
require 'active_sierra_models'
require 'axlsx'
require 'sqlite3'

module SerialList
  def self.all_potential_serial_orders
    (serial_orders_by_status_code + (serial_orders_by_funds)).uniq
  end

  def self.serial_orders_by_status_code
    OrderView.where(
      order_status_code: serial_status_codes,
      ocode1: ["u", "h"] ## Jurisdiction 
    )
  end

  def self.get_codes_from_funds
    ## We're using these to look up orders via OrderRecordCmf, which needs the codes as 5-digit strings
    (FundMaster.where(code: all_fund_list).collect { |f| "%05d" % f.code_num}).uniq
  end

  def self.serial_orders_by_funds
    ## This catches all the orders against serial funds with a monograph status code
    OrderView.joins(:order_record_cmf).where(
      order_record_cmf: { fund_code: get_codes_from_funds }).where(
      order_status_code: monograph_status_codes,
      ocode1: jurisdiction_codes
    )
  end

  def self.serial_status_codes
    ["c", "f", "d", "e", "g"]
  end

  def self.monograph_status_codes
    ["a", "o", "q"]
  end

  def self.jurisdiction_codes
    ["u", "h"]
  end

  def self.all_fund_list(fund_array = %w[ sbind scdnt scdrm scont sghum slanf smemb sp&pa sfren sgper slang sprof sspan ssocw scrce scrcl seduc slref senvi sgeog sgeol smath sphys sdaap sbiol schem sarch sdocs sccm sgerm sslav scrmj shums spsyc ssoci scoma sengl sling sthtr safro santh sasia shist sjuda slata sphil spols swoms scas scomp sengr sbusa secon halhs hhgms hhgps hhgps hhlos hhrcs hngs hnles hphs hygs hyss ysems ysmgs yells yoess ylecs yurbs ybots ydays yters ])
    ## Returns all funds unless another array is passed
    fund_array
  end
end

class List
  def initialize(orders)
    @order_views = orders
    @holdings = Holdings
  end

  def worksheet
    Axlsx::Package.new do |p|
      p.use_shared_strings = true ## This supports line breaks, per discussion here: https://github.com/randym/axlsx/issues/252
      p.workbook.add_worksheet(name: "Serial_orders") do |sheet|
        add_header(sheet)
        add_rows(sheet)
      end
      p.serialize('example.xlsx')
    end
  end

  def add_header(sheet)
    sheet.add_row(spreadsheet_keys(@order_views.first))
  end

  def add_rows(sheet)
    @order_views.each do |order_view|
      next unless order_view.record_metadata.deletion_date_gmt.nil?
      sheet.add_row(spreadsheet_values(order_view))
    end
  end
  
  def spreadsheet_keys(order_view)
    spreadsheet_mapping(order_view).collect { |row| row.keys[0]}
  end

  def spreadsheet_values(order_view)
    spreadsheet_mapping(order_view).collect { |row| row.values[0] }
  end

  def spreadsheet_mapping(order_view)
    [ 
      { order_number: "o#{order_view.record_num}a" },
      { title: order_view.bib_view.title },
      { issn1: issn_scan(order_view)[0] },
      { issn2: issn_scan(order_view)[1] },
      { online_access: holdings_array(order_view).join("\n") },
      { format: order_view.material_type_code },
      { fund: order_view.order_record_cmf.fund },
      { vendor: order_view.vendor_record_code },
      { acqusition_type: order_view.acq_type_code },
      { split?: order_view.receiving_action_code == "p" },
      { "FY#{fiscal_year(0)}".to_s => payment_total_for_fiscal_year(0, order_view) },
      { "FY#{fiscal_year(1)}".to_s => payment_total_for_fiscal_year(1, order_view) },
      { "FY#{fiscal_year(2)}".to_s => payment_total_for_fiscal_year(2, order_view) },
      { "FY#{fiscal_year(3)}".to_s => payment_total_for_fiscal_year(3, order_view) },
      { "FY#{fiscal_year(4)}".to_s => payment_total_for_fiscal_year(4, order_view) }
    ]
  end

  def issn_scan(order_view) 
    issn_marc_field = order_view.bib_view.varfield_views.marc_tag("022")
    return [ nil , nil ] if issn_marc_field.empty?
    issns = issn_marc_field.first.field_content.scan(/\d\d\d\d-\d\d\d[\dXx]/)
    return [issns[0], issns[1]] if issns.length == 2
    return [issns[0], nil ] if issns.length == 1
    [ nil , nil ]
  end

  def holdings_array(order_view)
    title_holdings = []
    issn_scan(order_view).each { |issn| next if issn.nil?;  title_holdings.concat(holdings_for(issn)) }
    title_holdings.uniq
  end

  def holdings_for(issn)
    holdings_text = []
    @holdings.where(issn: issn).each do |holding| 
      next if holding.nil?
      holdings_text.push("#{holding.startdate}-#{holding.enddate} | #{holding.resource}")
    end
    @holdings.where(eissn: issn).each do |holding|
      next if holding.nil?
      holdings_text.push("#{holding.startdate}-#{holding.enddate} | #{holding.resource}")
    end
    holdings_text
  end

  def fiscal_year(offset)
    ## offset retrieves previous fiscal years
    todays_month = Date.today.month
    todays_year = Date.today.year - offset
    return todays_year unless todays_month > 6
    return today_year + 1
  end

  def payment_total_for_fiscal_year(offset, order_view)
    year = fiscal_year(offset)
    fiscal_year_begin = Date.new(year - 1,7,1)
    fiscal_year_end = Date.new(year,6,30)
    payment_records = order_view.order_record_paids.where(paid_date_gmt: fiscal_year_begin..fiscal_year_end)
    (payment_records.collect { |payment| payment.paid_amount }).reduce(:+)
  end
end

class Holdings < ActiveRecord::Base
  ActiveRecord::Base.establish_connection(
    :adapter => "sqlite3",
    :database  => "holdings.db"
  )
end

l = List.new(SerialList.all_potential_serial_orders)
l.worksheet
