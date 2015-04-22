require 'bundler/setup'
require 'active_sierra_models'
require 'clipboard'

## Looking up information about orders:

# Copy your truncated order numbers into a comma delimited list on your clipboard
os = (Clipboard.paste).split(",")

order_objects = os.collect{ |o| OrderView.find_by_record_num(o) }

output = order_objects.collect do |o|
     puts o.record_num
     bib = o.bib_view
     ## Add other fields as needed
     title = bib.title
     puts title
     payment = o.order_record_paids.order(:paid_date_gmt).last.paid_amount unless o.order_record_paids.length == 0
     fund = o.order_record_cmf.fund
     format = o.material_type_code
     status = o.order_status_code
     acq_type = o.acq_type_code
     "#{title}\t#{payment}\t#{fund}\t#{format}\t#{status}\t#{acq_type}"
end

Clipboard.copy(output.join("\n"))
