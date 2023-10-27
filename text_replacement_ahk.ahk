#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
;trying to figure out how to put variables in strings and whatever
;:B0:mh::
;    Input, outp, V L5, {space}
;	Men := SubStr(outp, 1)
;	Hours := SubStr(outp, 2)
;   	int := %Men%
;	ManHours := int * 1000
;	numberOfBackSpaces:=strlen(outp) + 5
;	SendInput,{Backspace %numberOfBackSpaces%}{raw} %outp%%Men%%Hours%%ManHours%
;return
; ^ = ctrl, ! = alt, + = shift

;:*B0:wanted::
;	Input, name, V, {Enter}	; The name is stored in variable name, the option V makes whatever you type Visible, {enter} make enter complete the input
;	numberOfBackSpaces:=strlen(name) + 6 + 1 	; calculate the number of backspaces needed, backspace out the name, the word wanted (6), and the endkey enter (1)
;	SendInput,{Backspace %numberOfBackSpaces%}{raw}!setwantedlvl %name%,1
;return

;logs in quickbooks
^j::
Send, Roofing2 {Enter}
return

;switches quickbooks companies
^+j::
Send, !{f}{p}{down}{enter}
Sleep 10000
Send, Roofing2 {Enter}
return

;adds a line for a change order and the new contract amount for progress invoicing
^9::
Sleep 200
Send, L&M{Tab}+{End 2}{Delete}Change Order{Tab 2}0{Tab 2}L&M{Tab}+{End 2}{Delete}New contract amount{Tab 2}0{Up}+{Tab}
;Sleep 200
;SendEvent {Ctrl Up}
return



;sends tab 7 times from the last thing i usually enter in the fields of a bill and the job it's associated with
^e::
Send, {Tab 7}
return

!e::
Send, {Tab 4}
return

^l::
Send, {Enter 9}
return

;opens scanner
^+s::
Run, "C:\Program Files (x86)\Epson Software\Epson ScanSmart\ScanSmart.exe"
return

;opens programs i normally use on startup
^!s::
Run, C:\Program Files (x86)\Google\Chrome\Application\chrome.exe
Run, C:\Program Files (x86)\Intuit\QuickBooks 2019\QBW32Pro.exe
return

;names scanned documents from the draft file
!r::
Send, {F2}
Sleep 100
Send, ^{C}
ClipWait, 2
Sleep 1200
if WinExist("Epson ScanSmart")
    WinActivate
    WinWaitActive 
Send, {Enter}
Sleep 1000
Send, ^v
Sleep 1000
Send, {Enter}
SetTitleMatchMode, 2 
	IfWinExist, Google Chrome 
	WinActivate, Google Chrome 
	WinWaitActive, Google Chrome
Sleep 50
Sleep 500
;Send, {C}
;Sleep 1000
;Send, {Tab}
;Sleep 100
;Send, ^v
;Sleep 100
;Send, {Tab 4}
;Sleep 100
;Send, {Space}
;Sleep 1000
;Send, +{Tab 2}
;Send, {Home}
;Sleep 200
;Send, {Enter}
;Sleep 300
;SendInput Please see our attached proposal.{Enter}Thank you,{Enter}Michael Wormley{Enter}WBR Roofing{Enter}​O: 847-487-8787​{Enter}​wbrroof@aol.com
;Send, +{Tab 2}
return

;opens proposals
^+p::
Run, "\\WBR\shared\Proposals"
Return

;opens aia
^+g::
Run, "\\WBR\data\shared\G702 & G703 Forms"
Return

;opens blank proposal
^+b::
Run, "\\WBR\shared\Proposals\Blank.docx"
Return

;opens manual
^+m::
Run, "C:\Users\Michael\Desktop\Manual"
Return

;opens blank proposal
^+w::
Run, "C:\Users\Michael\Desktop\python-work"
Return

^+z::
    ; Google address search 
Send, +{Home}+{Up}^{c}
ClipWait, 3
    Sleep 200
    parameter = C:\Program Files (x86)\Google\Chrome\Application\chrome.exe https://www.google.com/search?q="%clipboard%"
    Run %parameter%
Return

;opens chrome
^g::
SetTitleMatchMode, 2 
	IfWinExist, Google Chrome 
	WinActivate, Google Chrome 
	WinWaitActive, Google Chrome
Return

^q::
SetTitleMatchMode, 2 
	IfWinExist, QuickBooks Desktop
	WinActivate, QuickBooks Desktop
	WinWaitActive, QuickBooks Desktop
	IfWinNotExist, QuickBooks Desktop
	Run, C:\Program Files (x86)\Intuit\QuickBooks 2019\QBW32Pro.exe
Return

^+e::
SetTitleMatchMode, 2 
	IfWinExist, CEnvelopes
	WinActivate, CEnvelopes
	Run, "C:\Users\Michael\Desktop\CEnvelopes"
Return

^!v::
SetTitleMatchMode, 2 
	IfWinExist, Visual Studio
	WinActivate, Visual Studio
	WinWaitActive, Visual Studio
Return

;bb billing autofill from hours sheets (BB side)
;you run this macro once you've input the billing name "WBR Roofing Company Inc.:Belle Tire W Chicago" and press tab
;so that the template is now highlighted
;in the excel worksheet you have the hours of the first job selected
^+l::
Send, {TAB 1}
Sleep 100
Send, {M}
Sleep 100
Send, {-}
Sleep 100
Send, {TAB 5}
Sleep 100
SetTitleMatchMode, 2
	IfWinExist, xlsx 
	WinActivate, xlsx
	WinWaitActive, xlsx
Sleep 1000
Send, ^c
ClipWait, 2
Sleep 300
Send, {TAB}
Sleep 100
Send, !{TAB}
Sleep 400
Send, ^v
Sleep 200
Send, !a
Sleep 4000
Send, !t
Return

;bb billing once spreadsheet of wbr invoices has been created from bb (WBR side)
;start by typing in the first job name under customer:job , then pressing tab, the first command is to uncheck billable
;make sure you have your wbr invoices to input report open! and highlight the first inv num
;^+t::
Send, {Space}
Send, +{TAB 10}
Sleep 500
Send, {M}
Sleep 200
Send, {-}
Sleep 200
Send, {TAB}
Sleep 200
;SetTitleMatchMode, 2
;	IfWinExist, Book3
;	WinActivate, Book3
;	WinWaitActive, Book3
SetTitleMatchMode, 2
	IfWinExist, xlsx
	WinActivate, xlsx
	WinWaitActive, xlsx
Sleep 1000
Send, ^c
ClipWait, 2
Sleep 300
Send, {TAB}
Sleep 200
Send, !{TAB}
Sleep 1000
Send, ^v
Sleep 300
Send, {TAB}
;SetTitleMatchMode, 2
;	IfWinExist, Book3
;	WinActivate, Book3
;	WinWaitActive, Book3
SetTitleMatchMode, 2
	IfWinExist, xlsx
	WinActivate, xlsx
	WinWaitActive, xlsx
Sleep 1000
Send, ^c
ClipWait, 2
Sleep 200
Send, {Left}
Sleep 300
Send, {Down}
Sleep 200
Send, !{TAB}
Sleep 1000
Send, ^v
Sleep 300
Send, !a
Sleep 2300
Send, !t
Sleep 1800
Send, {TAB 10}
Return

;date used for billing
^d::
Send, On 10/20/23 we 
Send, {Left 6}
return

;:*:@@::example@domain.com means when you type @@ it changes it to that email address
:*:ptf/::Photos to follow. {Enter}
:*:cuar/::Clean up and remove resulting debris.
:*:epdm::EPDM
:*X:hours/::Run pythonw.exe "C:\Users\Michael\Desktop\python-work\menandhours.pyw"
:*:e/::EPDM
:*:be/::black EPDM
:*:wt/::white TPO
:*:t/::TPO
:*:cd/::cleaned debris
:*:rw/::roofing work
:*:rw/::roofing work
:*:rec/::recaulked

:*:bl/::building
:*:pn/::Please note:	
:*:app/::approximately
:*:bal/::ballasted
:*:fa/::fully adhered
:*:totr/::2 layers total R=30 2.6” isocyanurate roof insulation
:*:fai/::furnish and install
:*:faid/::furnished and installed
:*:pfm::prefinished metal
:*:me/::metal
:*:pms::per manufacturer's specifications
:*:pusp::pop-up safety post
:*:rm/::roof membrane
:*:irm/::in roof membrane
:*:sa/::self-adhering
:*:vb/::vapor barrier
:*:rtu/::roof top unit
:*:rtuc/::roof top unit curb
:*:rtc/::rtu curb
:*:bd/::Balance due
:*:ssp/::soil stack pipe
:*:os/::open seam
:*:exos/::open seam in the roof membrane
:*:exoss/::open seams in the roof membrane
:*:itr/::in the roof membrane
:*:oc/::oiling canning
:*:so/::soffit
:*:cl/::come loose
:*:sof/::soffit
:*:de/::drip edge
:*:pw/::parapet wall
:*:archl/::architectural
:*:archt/::architect
:*:sh/::shingle
:*:kow/::kind of work
:*:ttp/::through the portal
:*:tht/::through the 
:*:ss/::standing seam
:*:iasc/::in a standard color
:*:aco/::^{b}Added Cost: $ 
:*:bp/::budget pricing
:*:demo/::demolition
:*:dec/::decking
:*:cp/::composite panel
:*:ps/::popped screws
:*:nw/::night work
:*:iipc/::is in poor condition
:*:rs/::roof screen
:*:tp/::tuckpointing
:*:ins/::insulation
:*:inst/::install
:*:instd/::install
:*:carp/::carpentry
:*:w/::wall
:*:se/::seal
:*:sed/::sealed
:*:bot/::by others
:*:bod/::build out
:*:pi/::pipe
:*:sl/::skylight
:*:mp/::metal panel
:*:thf/::that had failed
:*:wp/::wall panel
:*:cu/::coming through
:*:gal/::galvanized
:*:ptd/::paid to date
:*:ef/::exhaust fan
:*:efc/::exhaust fan curb
:*:rc/::roof top unit curb
:*:com/::composite
:*:fe/::front entrance
:*:md/::metal decking
:*:iaws/::ice and water shield
:*:cho/::Change Order
:*:chor/::Change Order Request
:*:psa/::Please see attached.
:*:psap/::Hello,{Enter}Please see our attached proposal.{Enter}Thank you,{Enter 2}Michael Wormley{Enter}WBR Roofing{Enter}​O: 847-487-8787​{Enter}​wbrroof@aol.com
:*:mw/::{Enter 2}Michael Wormley{Enter}WBR Roofing{Enter}​O: 847-487-8787​{Enter}​wbrroof@aol.com
:*:tymw/::Thank you,{Enter 2}Michael Wormley{Enter}WBR Roofing{Enter}​O: 847-487-8787​{Enter}​wbrroof@aol.com
:*:war/::William A. Randolph
:*:pac/::Petersen Aluminum Corp.
:*:can/::canopy
:*:cans/::canopies
:*:tec/::trash enclosure coping
:*:ds/::downspout
:*:dp/::drain pipe
:*:gael/::gas and electric lines
:*:gl/::gas line
:*:eb/::electric box
:*:insp/::inspection
:*:ops/::open seam
:*:rel/::reported leak
:*:rol/::roof leak
:*:gs/::gravel stop
:*:gb/::gravel ballast
:*:ymw/::You may want
:*:rh/::roof hatch
:*:er/::existing roof
:*:PSTC::Note: Pricing is subject to change due to material supply costs. This price valid through
:*:1l/::1-layer
:*:dd/::DensDeck
:*:mb/::modified bitumen
:*:ra/::roof area
:*:gr/::green roof
:*:pow/::ponding water
:*:lam/::Labor and material
:*:pc/::per contract
:*:cpc/::completed per contract
:*:exlam/::Labor and material for roofing work completed per contract.
:*:repp/::Hello,{Enter 2}Please see attached photos, thank you.{Enter 2}Michael Wormley{Enter}WBR Roofing{Enter}​O: 847-487-8787​{Enter}​wbrroof@aol.com
:*:wn/::wood nailer
:*:mem/::membrane
:*:nsp/::National Shopping Plazas
:*:ty/::Thank you
:*:cs/::chimney saddle
:*:ex/::existing
:*:sm/::sheet metal
:*:ew/::extra work
:*:bo/::Build Out
:*:dop/::Date of Plans: 
:*:pb/::pipe boot
:*:pch/::pipe chase
:*:cb/::collector box
:*:wf/::wall flashing
:*:cf/::corner flashing
:*:uf/::counter flashing
:*:wr/::We repaired
:*:ws/::We sealed
:*:wfo/::We found
:*:24ga/::24 ga prefinished metal
:*:prev/::Previously billed
:*:pr/::Previously retained
:*:ct/::Completed through
:*:nca/::New contract amount
:*:sr/::safety rail
:*:pen/::penetration
:*:peo/::per our existing
:*:poe/::per our existing
:*:manu/::Manufacturer’s Warranty
:*:wty/::Warranty
:*:pt/::Per Tom:
:*:tc/::Tim Cote
:*:mal/::Manufacturer's material and labor warranty
:*:cond/::condenser
:*:cov/::cover board
:*:lm/::Landmark
:*:fl/::flashing
:*:refl/::reflash
:*:refld/::reflashed
:*:f/::flashing
:*:fld/::flashed
:*:fls/::flash
:*:cfl/::counterflashing
:*:cfl/::counterflashed
:*:cfs/::counterflash
:*:sq/::sq. ft.
:*:subsworn/::Subcricbed and sworn to before me this *** day of ***
:*:rep/::repair
:*:repd/::repaired
:*:repl/::replace
:*:ti/::tie into
:*:fas/::fascia
:*:rem/::remove
:*:gu/::gutter
:*:ap/::apron
:*:gua/::gutter apron
:*:an/::as necessary
:*:main/::maintenance
:*:uc/::Upon Completion
:*:5d/::50% Down, 50% Upon Completion
:*:po/::PAYOUTS
:*:res/::residence
:*:0b/::060 black EPDM
:*:sp/::single ply
:*:mul/::multiple
:*:pl/::plywood
:*:ply/::plywood
:*:val/::valley
:*:pp/::pitch pan
:*:add/::additional
:*:ad/::access door
:*:ca/::caulk
:*:cad/::caulked
:*:inv/::invoice
:*:chim/::chimney
:*:rein/::reinstall
:*:sms/::sheet metal supply
:*:cov/::cover board
:*:lr/::leak repair
:*:ej/::expansion joint
:*:tap/::tapered
:*:scup/::scupper
:*:sb/::scupper box
:*:sec/::section
:*:rp/::revised proposal
:*:r/::roof
:*:g/::gravel
:*:p/::patch
:*:serv/::service side
:*:por/::portal
:*:nec/::necessary
:*:ref/::refasten
:*:hd/::high density
:*:ew/::extra work
:*:hs/::heat stack
:*:pyr/::per your request
:*:pomc/::part of maintenance contract
:*:pomw/::part of maintenance work
:*:mar/::maintenance and repair work
:*:timw/::This is maintenance work
:*:ncuw/::not covered under warranty
:*:ri/::roof inspection
:*:rer/::reroof
:*:wy/::warranty
:*:nwy/::non-warranty
:*:dam/::damage
:*:hv/::HVAC
:*:JM/::Johns Manville
:*:ln/::ln feet
:*:exh/::exhaust
:*:c/::curb
:*:rev/::revised
:*:nc/::no charge
:*:el/::electric line
:*:gad/::gutters and downspouts
:*:gads/::gutters and downspouts
:*:p/::portal
:*:att/::attach
:*:ph/::photo
:*:wa/::We also
:*:ward/::We also repaired
:*:wer/::we repaired
:*:fbp/::Firestone Building Products
:*:fs/::Firestone
:*:aso/::as shown on roof plan
:*:tme/::to match existing
:*:mat/::material
:*:yb/::your building of approx
:*:jai/::jgalbraith@gajohnson.com {space} igalbraith@gajohnson.com {space}
:*:dl/::Duro-Last
:*:tt/::Please print the following pages: {Enter} ^{v} {Enter} Thank you.{Enter}Michael Wormley{Enter}WBR Roofing{Enter}​O: 847-487-8787​{Enter}​wbrroof@aol.com
:*:sf/::stepflashing
:*:sc/::schedule
:*:m/::metal
:*:gem/::Gemco Supply
:*:abc/::ABC Supply
:*:wbr/::WBR Roofing Company Inc.
:*:25084/::25084 W Old Rand Rd, Wauconda, IL 60084
:*:pa/::please advise
:*:starb/::Starbucks
:*:ml/::multiple locations
:*:lamf/::Labor and material for roofing work per contract.
:*:ab/::Alt Bid
:*:adb/::Additional Bid
:*:4k/::4-knob
:*:ac/::^{b}Added Cost: $
:*:em/::7/31/23
:*:eom/::7/31/23
:*:ey/::12/31/22
:*:eoy/::12/31/22
:*:anyq/::Let me know if you have any questions.

:*:rr/::roof repair
:*:fm/::field membrane
:*:arch/::architect
:*:pig/::Price is good through
:*:ww/::walkway
:*:wwa/::Walkway as shown on roof plan.
:*:nm/::(needs measurement)
:*:nt/::(needs time)
:*:me/::match existing
:*:aw/::additional work
:*:rtw/::roof to wall
:*:bw/::brick wall
:*:op/::old patch
:*:ge/::gutter edge
:*:pro/::proposal
:*:strh/::2'6" x 3'
:*:yn/::You'll need an HVAC contractor to do this work.
:*:ofs/::open field seam
:*:l/::leak
:*:con/::contractor
:*:cons/::construction
:*:h/::hole
:*:ha/::hatch
:*:v/::vent
:*:dw/::duct work
:*:waf/::We also found
:*:cor/::corrugated
:*:ctap/::coming through a portal
:*:ctu/::coming through the portals
:*:rpo/::Repair Photos
:*:cac/::caulked and clamped
:*:hof/::Hofstadter
:*:Exc/::^{b}Exclude: ^{b}
:*:duc/::^{b}DUE UPON COMPLETION ^{b}
:*:Excl/::^{b}Exclude: ^{b}
:*:rd/::Retention due
:*:sv/::smoke vent
:*:mez/::mezzanine
:*:rl/::roof leak
:*:nb/::Not bidding
:*:rel/::reported leak
:*:mpg/::Midwest Property Group
:*:1y/::WBR 1-Year labor guarantee.
:*:hw/::Hello, world!
:*:wwb/::We won't be bidding this one.
:*:wwr/::which we repaired
:*:wwp/::which we patched
:*:wc/::reported leak wasn't coming from the roof
:*:sos/::soil stack
:*:to/::tear off
:*:dtd/::down to deck
:*:u/::portal
:*:cn/::CITATION NEEDED
:*:ai/::anti-intellectualism
:*:met/::(Marques et al., 2017)
:*:too/::2’6”x3’
:*:tmec/::to meet energy code
:*:itrm/::in the roof membrane
:*:gt/::going through




;zips
:*:lf/::Lake Forest, IL 60045
:*:countryside/::Countryside, IL 60525
:*:rlb/::Round Lake Beach, IL 60073
:*:lz/::Lake Zurich, IL 60047
:*:rose/::Rosemont, IL 60018
:*:niles/::Niles, IL 60714
:*:huntley/::Huntley, IL 60142
:*:carpentersville/::Carpentersville, IL 60110
:*:elb/::Elburn, IL 60119
:*:stc/::St Charles, IL 60174
:*:jol/::Joliet, IL 60435
:*:nv/::Naperville, IL 60540
:*:Naperville/::Naperville, IL 60540
:*:chi/::Chicago, IL 60618
:*:northlake/::Northlake, IL 60164
:*:skokie/::Skokie, IL 60077
:*:highland/::Highland Park, IL 60035
:*:bensen/::Bensenville, IL 60106
:*:buffg/::Buffalo Grove, IL 60089		
:*:gonq/::Algonquin, IL 60102
:*:schiller/::Schiller Park, IL 60176
:*:tinpark/::Tinley Park, IL 60477
:*:lincoln/::Lincolnshire, IL 60069
:*:vh/::Vernon Hills, IL 60061
:*:lv/::Libertyville, IL 60048
:*:hp/::Highland Park, IL 60035
:*:Glen Ellyn/::Glen Ellyn, IL 60137
:*:North Aurora/::North Aurora, IL 60542
:*:Oak Brook/::Oak Brook, IL 60523
:*:Lockport/::Lockport, IL 60491
:*:Palatine/::Palatine, IL 60074
:*:Glendale Heights/::Glendale Heights, IL 60139
:*:Glendale/::Glendale Heights, IL 60139
:*:Elmhurst/::Elmhurst, IL 60126
:*:Downers Grove/::Downers Grove, IL 60515
:*:Mundelein/::Mundelein, IL 60060
:*:Elburn/::Elburn, IL 60119
:*:Lake Bluff/::Lake Bluff, IL 60044
:*:Des Plaines/::Des Plaines, IL 60018
:*:Cicero/::Cicero, IL 60804
:*:Crystal Lake/::Crystal Lake, IL 60014
:*:Lake in the Hills/::Lake in the Hills, IL 60156
:*:Downers Grove/::Downers Grove, IL 60515
:*:Richton Park/::Richton Park, IL 60471
:*:Yorkville/::Yorkville, IL 60560
:*:Huntley/::Huntley, IL 60142
:*:Joliet/::Joliet, IL 60563
:*:Lindenhurst/::Lindenhurst, IL 60046
:*:Barrington Hills/::Barrington Hills, IL 60010
:*:Schaumburg/::Schaumburg, IL 60173
:*:Highland Park/::Highland Park, IL 60035
:*:Bridgeview/::Bridgeview, IL 60455
:*:Crestwood/::Crestwood, IL 60445
:*:Spring Grove/::Spring Grove, IL 60081
:*:Rosemont/::Rosemont, IL 60018
:*:Evanston/::Evanston, IL 60202
:*:Morton Grove/::Morton Grove, IL 60053
:*:mg/::Morton Grove, IL 60053
:*:Lake in the Hills/::Lake in the Hills, IL 60156
:*:LITH/::Lake in the Hills, IL 60156
:*:geneva/::Geneva, IL 60134
:*:hinsdale/::Hinsdale, IL 60521
:*:Elgin/::Elgin, IL 60124
:*:Deer Park/::Deer Park, IL 60010
:*:Cary/::Cary, IL 60013
:*:Lincolnshire/::Lincolnshire, IL 60069
:*:Aurora/::Aurora, IL 60502
:*:arlington/::Arlington Heights, IL 60005
:*:ah/::Arlington Heights, IL 60005
:*:Wheaton/::Wheaton, IL 60187
:*:Batavia/::Batavia, IL 60510
:*:Calumet/::Calumet City, IL 60409
:*:Barrington/::Barrington, IL 60010
:*:Bar/::Barrington, IL 60010
:*:Northbrook/::Northbrook, IL 60062
:*:Kildeer/::Kildeer, IL 60047
:*:Northfield/::Northfield, IL 60093
:*:Wauconda/::Wauconda, IL 60084
:*:Antioch/::Antioch, IL 60002
:*:Libertyville/::Libertyville, IL 60048
:*:Grayslake/::Grayslake, IL 60030
:*:Deerfield/::Deerfield, IL 60015
:*:New Lenox/::New Lenox, IL 60451
:*:Park Ridge/::Park Ridge, IL 60068
:*:Libertyville/::Libertyville, IL 60048
:*:Algonquin/::Algonquin, IL 60102
:*:Rolling Meadows/::Rolling Meadows, IL 60008
:*:Grayslake/::Grayslake, IL 60030
:*:Lilse/::Lisle, IL 60532
:*:Orland Park/::Orland Park, IL 60462
:*:Hoffman Estates/::Hoffman Estates, IL 60192
:*:Carol Stream/::Carol Stream, IL 60188
:*:Grayslake/::Grayslake, IL 60030
:*:Vernon Hills/::Vernon Hills, IL 60061
:*:Bensenville/::Bensenville, IL 60106


;text replacement for commands
:*:slate/::+{End}{Delete}{Down}{End}+{Home}{Delete}{Down}{End}+{Home}{Delete}{Up 2}
;slate 2 deletes the arch and date of plans
:*:s2/::+{End}{Delete}{Down}{End}+{Home}{Delete}{Down}{End}+{Home}{Delete}{Down}{End}+{Home}{Delete}Architect: {Down}{End}+{Home}{Delete}Date of Plans: {Up 4}
;deletes just the address but not the title
:*:s3/::+{End}{Delete}{Down}{End}+{Home}{Delete}{Up 1}
;deletes the Date of Plans
:*:s4/::+{Home}{Delete}Date of Plans:{space} 


;pulls address info from proposals
^!a::
FileAppend, `n:*:COMPANY/::, C:\Users\Michael\Documents\companyaddresses.txt
send, +{Home}^{c}
Sleep 1000
ClipWait, 2
FileAppend, +{End}{Delete}%clipboard%{Down}{End}+{Home}, C:\Users\Michael\Documents\companyaddresses.txt
Send, {Down}{End}+{Home}^{c}
Sleep 1000
ClipWait, 2
FileAppend, %clipboard%{Down}{End}+{Home}, C:\Users\Michael\Documents\companyaddresses.txt
Send, {Down}{End}+{Home}^{c}
Sleep 1000
ClipWait, 2
FileAppend, %clipboard%{Down}{End}^+{Left 2}, C:\Users\Michael\Documents\companyaddresses.txt
Send, {Down}{End}{End}^+{Left 2}^{c}
Sleep 1000
ClipWait, 2
FileAppend, %clipboard%`n, C:\Users\Michael\Documents\companyaddresses.txt
;
;{Up 3}{End}{Space}^{v}{Down 4}{End}+{Home}^{c}{Up 4}{End}{Space}^{v}
;send, {Down 5}!{Tab}^{c}{Up5}{End}{Space}^{v}+{Home}^{x}
return




;Company Submissions (fills in address of companies 
:*:41 North/::+{End}{Delete}41 North Contractors{Down}{End}+{Home}4906 Main St Ste 102{Down}{End}+{Home}Lisle, IL 60532{Down}{End}^+{Left 2}DJ Tonyan
:*:raby/::+{End}{Delete}Raby Roofing, Inc. {Down}{End}+{Home}210 Industrial Dr{Down}{End}+{Home}Gilberts, IL 60136{Down}{End}^+{Left 2}Dave Raby
:*:healy/::+{End}{Delete}Healy Construction Services, Inc.{Down}{End}+{Home}14000 Keeler Ave{Down}{End}+{Home}Crestwood, IL 60418{Down}{End}^+{Left 2}Lynn Tabor
:*:delko/::+{End}{Delete}Delko Construction Co Inc.{Down}{End}+{Home}4849 N Milwaukee Ave #102{Down}{End}+{Home}Chicago, IL 60630{Down}{End}^+{Left 2}Laura Koulis
:*:dejames/::+{End}{Delete}DeJames Builders, Inc.{Down}{End}+{Home}1957 Quincy Court, Ste 101{Down}{End}+{Home}Glendale Heights, IL 60139{Down}{End}^+{Left 2}Pam Lendy
:*:walter/::+{End}{Delete}Walter Daniels Construction Co.{Down}{End}+{Home}6316 N Northwest Hwy{Down}{End}+{Home}Chicago, IL 60654{Down}{End}^+{Left 2}Robert Arnolde
:*:novak/::+{End}{Delete}Novak Construction{Down}{End}+{Home}3423 N Drake Ave{Down}{End}+{Home}Chicago, IL 60618{Down}{End}^+{Left 2}Mike Krzyston
:*:pmreal/::+{End}{Delete}PM Realty{Down}{End}+{Home}4333 S Pulaski{Down}{End}+{Home}Chicago, IL 60632{Down}{End}^+{Left 2}Eugene Grzynkowicz
:*:wilar/::+{End}{Delete}William A. Randolph, Inc.{Down}{End}+{Home}820 Lakeside Dr #3{Down}{End}+{Home}Gurnee, IL 60031{Down}{End}^+{Left 2}Peter Farquhar
:*:osman/::+{End}{Delete}Osman Construction{Down}{End}+{Home}70 W Seegers Rd{Down}{End}+{Home}Arlington Heights, IL 60005{Down}{End}^+{Left 2}Walter Johnson
:*:cote/::+{End}{Delete}Tim Cote, Inc.{Down}{End}+{Home}1075 Manito Trail{Down}{End}+{Home}Algonquin, IL 60102{Down}{End}^+{Left 2}Brian Adamczyk
:*:dzi/::+{End}{Delete}DZI Construction Services Inc.{Down}{End}+{Home}9675 Northwest Ct{Down}{End}+{Home}Village of Clarkston, MI 48346{Down}{End}^+{Left 2}David Huber
:*:shorewood/::+{End}{Delete}Shorewood Development Group{Down}{End}+{Home}790 Estate Dr Ste 200{Down}{End}+{Home}Deerfield, IL 60015{Down}{End}^+{Left 2}Dan Angspatt
:*:rcarl/::+{End}{Delete}R Carlson & Sons, Inc.{Down}{End}+{Home}19140 104th Ave{Down}{End}+{Home}Mokena, IL 60448{Down}{End}^+{Left 2}Nick Cannova
:*:vequity/::+{End}{Delete}Vequity Construction{Down}{End}+{Home}226 N Morgan St 400{Down}{End}+{Home}Chicago, IL 60607{Down}{End}^+{Left 2}Tim Leung
:*:mario/::+{End}{Delete}Mariottini Construction, Inc.{Down}{End}+{Home}445 West Kay Ave{Down}{End}+{Home}Addison, IL 60101{Down}{End}^+{Left 2}Jeff Mariottini
:*:rosewood/::+{End}{Delete}Rosewood Construction{Down}{End}+{Home}1300 Howard Street{Down}{End}+{Home}Elk Grove Village, IL 60007{Down}{End}^+{Left 2}Larry Prace
:*:gaj/::+{End}{Delete}G.A. Johnson & Son{Down}{End}+{Home}828 Foster St{Down}{End}+{Home}Evanston, IL 60201{Down}{End}^+{Left 2}Ian Galbraith
:*:gallant/::+{End}{Delete}Gallant Construction Company, Inc.{Down}{End}+{Home}345 Memorial Dr{Down}{End}+{Home}Crystal Lake, IL 60014{Down}{End}^+{Left 2}Wes Siete
:*:stasica/::+{End}{Delete}Stasica Construction{Down}{End}+{Home}465 Spring Rd{Down}{End}+{Home}Elmhurst, IL 60126{Down}{End}^+{Left 2}Dominick Burke
:*:builtech/::+{End}{Delete}Builtech Services, LLC{Down}{End}+{Home}425 N Martingale Road 1050{Down}{End}+{Home}Schaumburg, IL 60173{Down}{End}^+{Left 2}Vincent Searson
:*:jg/::+{End}{Delete}J.G. Morris, Jr. Construction & Design{Down}{End}+{Home}22021 Commerce Dr{Down}{End}+{Home}Woodhaven, MI 48183{Down}{End}^+{Left 2}Scot Brand
:*:jgm/::+{End}{Delete}J.G. Morris, Jr. Construction & Design{Down}{End}+{Home}22021 Commerce Dr{Down}{End}+{Home}Woodhaven, MI 48183{Down}{End}^+{Left 2}Scot Brand
:*:morris/::+{End}{Delete}J.G. Morris, Jr. Construction & Design{Down}{End}+{Home}22021 Commerce Dr{Down}{End}+{Home}Woodhaven, MI 48183{Down}{End}^+{Left 2}Scot Brand
:*:mcnelly/::+{End}{Delete}McNelly Services, Inc.{Down}{End}+{Home}225 Oakwood Rd{Down}{End}+{Home}Lake Zurich, IL 60047{Down}{End}^+{Left 2}Dan McNelly
:*:cahill/::+{End}{Delete}Cahill Construction{Down}{End}+{Home}5233 Bethel Centre Mall{Down}{End}+{Home}Columbus, OH 43220{Down}{End}^+{Left 2}Judd Gerlach
:*:kps/::+{End}{Delete}KPS Commercial Construction{Down}{End}+{Home}1318 E 236th Street{Down}{End}+{Home}Arcadia, IN 46030{Down}{End}^+{Left 2}Estimating
:*:mrg/::+{End}{Delete}MRG Construction Corp.{Down}{End}+{Home}200 N LaSalle St Ste 2350{Down}{End}+{Home}Chicago, IL 60606{Down}{End}^+{Left 2}Paul Colletti
:*:capitol/::+{End}{Delete}Capitol Construction{Down}{End}+{Home}1050 W Rt 126{Down}{End}+{Home}Plainfield, IL 60544{Down}{End}^+{Left 2}Justin Bennett
:*:pdb/::+{End}{Delete}Pauley Drive Buildings, LLC{Down}{End}+{Home}956 S Bartlett Rd Suite 116{Down}{End}+{Home}Bartlett, IL 60103{Down}{End}^+{Left 2}Suzanne Maffia
:*:raff/::+{End}{Delete}Raffin Construction{Down}{End}+{Home}744 E 113th St{Down}{End}+{Home}Chicago, IL 60628{Down}{End}^+{Left 2}Mike Raffin
:*:sachse/::+{End}{Delete}Sachse Construction{Down}{End}+{Home}3663 Woodward Ave{Down}{End}+{Home}Detroit, MI 48201{Down}{End}^+{Left 2}Antoinette Miller
:*:mcp/::+{End}{Delete}Midwest Construction Partners, Inc.{Down}{End}+{Home}1300 Woodfield Dr{Down}{End}+{Home}Schaumburg, IL 60173{Down}{End}^+{Left 2}Estimating
:*:falcon/::+{End}{Delete}Falcon Construction{Down}{End}+{Home}Falcon Construction{Down}{End}+{Home}Falcon Construction{Down}{End}^+{Left 2}Jessie Lozano
:*:delauter/::+{End}{Delete}DeLauter, Inc.{Down}{End}+{Home}500 Coventry Ln, Ste 280{Down}{End}+{Home}Crystal Lake, IL 60014{Down}{End}^+{Left 2}Steve Tarzian
:*:Englewood/::+{End}{Delete}Englewood Construction{Down}{End}+{Home}80 Main St{Down}{End}+{Home}Lemont, IL 60439{Down}{End}^+{Left 2}Lupe Cortez
:*:Jirsa/::+{End}{Delete}Jirsa Construction Company{Down}{End}+{Home}806 Penny Ave Rt 68{Down}{End}+{Home}East Dundee, IL 60118{Down}{End}^+{Left 2}Eric Peca
:*:Divita/::+{End}{Delete}J. Divita & Associates{Down}{End}+{Home}250 Telser Rd{Down}{End}+{Home}Lake Zurich, IL 60047{Down}{End}^+{Left 2}John Divita
:*:Horizon/::+{End}{Delete}Horizon Management{Down}{End}+{Home}1540 E Dundee Rd Suite 240{Down}{End}+{Home}Palatine, IL 60074{Down}{End}^+{Left 2}Lisa Lee
:*:Schall/::+{End}{Delete}Schall Development {Down}{End}+{Home}26774 Longmeadow Cir{Down}{End}+{Home}Mundelein, IL 60060{Down}{End}^+{Left 2}John Kerrigan
:*:Sterling/::+{End}{Delete}Sterling Renaissance, Inc.{Down}{End}+{Home}430 East IL Rte 22{Down}{End}+{Home}Lake Zurich, IL 60047{Down}{End}^+{Left 2}Bruce Sterling
:*:Gleeson/::+{End}{Delete}C.E. Gleeson Constructors Inc.{Down}{End}+{Home}984 Livernois Road{Down}{End}+{Home}Troy, MI 48083{Down}{End}^+{Left 2}Janelle Wrublewski
:*:Hanna/::+{End}{Delete}Hanna Design Group, Inc.{Down}{End}+{Home}650 E Algonquin Rd Ste 405{Down}{End}+{Home}Schaumburg, IL 60173{Down}{End}^+{Left 2}Gary Knepper
:*:HDG/::+{End}{Delete}Hanna Design Group, Inc.{Down}{End}+{Home}650 E Algonquin Rd Ste 405{Down}{End}+{Home}Schaumburg, IL 60173{Down}{End}^+{Left 2}Gary Knepper
:*:Valenti/::+{End}{Delete}Valenti Builders, Inc.{Down}{End}+{Home}333 S Wabash Ave{Down}{End}+{Home}Chicago, IL 60604{Down}{End}^+{Left 2}Jack Scapin
:*:gh/::+{End}{Delete}G&H Developers Corp.{Down}{End}+{Home}200 W Madison Ste 4200{Down}{End}+{Home}Chicago, IL 60606{Down}{End}^+{Left 2}Kevin Kissman
:*:masonry/::+{End}{Delete}Rosemont Masonry{Down}{End}+{Home}9575 W Higgins Rd 902{Down}{End}+{Home}Rosemont, IL 60018{Down}{End}^+{Left 2}Dan Degen
:*:weiss/::+{End}{Delete}Weiss Construction{Down}{End}+{Home}3011 W 183rd Street, Suite 157{Down}{End}+{Home}Homewood, IL 60430{Down}{End}^+{Left 2}Russell Weissenhofer
:*:Mega/::+{End}{Delete}Mega Properties Inc. El Bawadi Restaurant Bridgeview{Down}{End}+{Home}1699 East Chicago St{Down}{End}+{Home}Elgin, IL 60120{Down}{End}^+{Left 2}Joe Dubs
:*:Dubs/::+{End}{Delete}The Dubs Company{Down}{End}+{Home}1699 East Chicago St{Down}{End}+{Home}Elgin, IL 60120{Down}{End}^+{Left 2}Joe Dubs
:*:Resa/::+{End}{Delete}RESA Construction{Down}{End}+{Home}711 N Elmhurst Rd{Down}{End}+{Home}Prospect Heights, IL 60070{Down}{End}^+{Left 2}Brandon Amodeo
:*:Trailblazer/::+{End}{Delete}Trailblazer Construction and Restoration LLC{Down}{End}+{Home}27031 South Sylvan Ln{Down}{End}+{Home}Monee, IL 60484{Down}{End}^+{Left 2}Ken Becker
:*:Loberg/::+{End}{Delete}Loberg Construction{Down}{End}+{Home}311 E. Illinois Ave{Down}{End}+{Home}Palatine, IL 60067{Down}{End}^+{Left 2}Jessica Takeda
:*:vpm/::+{End}{Delete}VP Mechanical{Down}{End}+{Home}1760 Britannia Dr, Ste 2{Down}{End}+{Home}Elgin, IL 60124{Down}{End}^+{Left 2}Dan Freeman
:*:degen/::+{End}{Delete}Degen and Rosato Construction Co{Down}{End}+{Home}9575 W Higgins Rd 902{Down}{End}+{Home}Rosemont, IL 60018{Down}{End}^+{Left 2}Ray Rosato
:*:dinaso/::+{End}{Delete}DiNaso & Sons Construction Co. Inc.{Down}{End}+{Home}9910 W 190th Street Suite A{Down}{End}+{Home}Mokena, IL 60448{Down}{End}^+{Left 2}Chuck DiNaso
:*:schwabe/::+{End}{Delete}Peter Schwabe Inc.{Down}{End}+{Home}13890 Bishop Drive, Suite 100{Down}{End}+{Home}Broookfield WI 53005{Down}{End}^+{Left 2}Kristian Nielsen
:*:innovative/::+{End}{Delete}Innovative Construction Solutions{Down}{End}+{Home}21675 Gateway Rd{Down}{End}+{Home}Brookfield, WI 53045{Down}{End}^+{Left 2}Lynn Key
:*:ics/::+{End}{Delete}Innovative Construction Solutions{Down}{End}+{Home}21675 Gateway Rd{Down}{End}+{Home}Brookfield, WI 53045{Down}{End}^+{Left 2}Lynn Key
:*:41 North/::+{End}{Delete}41 North Contractors{Down}{End}+{Home}4906 Main St Ste 102{Down}{End}+{Home}Lisle, IL 60532{Down}{End}^+{Left 2}DJ Tonyan
:*:Horizon Retail/::+{End}{Delete}Horizon Retail{Down}{End}+{Home}9999 E Exploration Ct{Down}{End}+{Home}Sturtevant, WI 53177{Down}{End}^+{Left 2}Judi Reid-Hart

;Arch
:*:chip/::Chipman Design



;attaches proposal once you're in gmail and have opened and added the contact
^+a::
Sleep 100
Send, {Tab}
Sleep 100
Send, %clipboard%
Sleep 300
Send, {Tab 4}
Sleep 200
Send, {Space}
Sleep 1500
Send, +{Tab}
Sleep 500
Send, +{Tab}
Sleep 200
Send, {Home}
Sleep 300
Send, {Enter}
Sleep 2000
Sleep 100
Send, Hello, {Enter 2}Please see our attached proposal.{Enter 2}Thank you,{Enter 2}Michael Wormley{Enter}WBR Roofing{Enter}​O: 847-487-8787​{Enter}​wbrroof@aol.com
;Send, Hello, {Enter 2}Please see attached billing.{Enter 2}Thank you,{Enter 2}Michael Wormley{Enter}WBR Roofing{Enter}​O: 847-487-8787​{Enter}​wbrroof@aol.com
Return

^!q::
send, ^{c}
Sleep 500
ClipWait, 2
Send, {tab 30}
send, {space}
send, +{End}^{v}
send, {tab}^{v}
send, {End}{Enter}


!y::
send, +{Home}^{c}
Sleep 1000
ClipWait, 2
FileAppend, :*:archi/::%clipboard%`n, C:\Users\Michael\Documents\architects.txt
return

!z::
send, +{Home}^{c}
Sleep 1000
ClipWait, 2
FileAppend, :*:t/::%clipboard%`n, C:\Users\Michael\Documents\zip.txt
return

;attaches most recent scan 
^+q::
Send, {Tab 3}
Sleep 200
Send, {Space}
Sleep 1000
Send, +{Tab 2}
Sleep 100
Send, {Home}
Sleep 200
Send, {Enter}
Sleep 2000
Send, Hello, {Enter 2}Please see attached.{Enter 2}Thank you,{Enter 2}Michael Wormley{Enter}WBR Roofing{Enter}​O: 847-487-8787​{Enter}​wbrroof@aol.com
Return

^+i::
send, {f}
Sleep 1000
Send, +{Tab 3}
Sleep 500
Send, {Space}
Sleep 100
Send, {Down 3}
Sleep 50
Send, {Enter}
Sleep 1000
Send, {Delete}Photos for Invoice {Tab}
Sleep 100
Send, +{PgDn 3}{Delete}
Send,Hello,{Enter 2}Please see attached photos, thank you.{Enter 2}Michael Wormley{Enter}WBR Roofing{Enter}​O: 847-487-8787​{Enter}​wbrroof@aol.com
Sleep 100
Sleep 100
Send, +{Tab 2}
Return



:*:bld/::building
:*:fab/::fabricate
:*:test/::test
:*:fab/::fabricate
:*:clip/::starter clip
:*:a./::git add .
:*:c./::git commit -m "
:*:p./::git push origin main

:*:narl/::It is not a roof leak
:*:loc/::locations
:*:si/::stripped in


; This hotkey (Windows+J) triggers a PowerShell script that prompts the user for a search term
; and then opens the specified folder in File Explorer with the search term already filled in, 
; displaying the search results.
^!p:: 
Run, powershell -ExecutionPolicy Bypass -File "C:\Users\Michael\Desktop\python-work\folderSearch.ps1"
return




^!+s::
Run, "proposal_scan.py"
return

^!+a::
Run, "AIA_scan.py"
return
:*:port/::Portillo's
:*:nrl/::not a roof leak
:*:iigc/::is in good condition
:*:iibc/::is in bad condition
:*:iiec/::is in excellent condition
:*:cc/::coping cap
:*:d/::drain