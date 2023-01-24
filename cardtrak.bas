Attribute VB_Name = "Module3"
Option Explicit

' This module contains the cardtrak summary functions

Public Function is_cardtrak_transaction(s As String) As Boolean
  is_cardtrak_transaction = False
  If (UCase(Mid(s, 2, 2)) = "CT") Then
    If (Val(Mid(s, 4, 2)) > 0) Then is_cardtrak_transaction = True
  End If
End Function


Public Sub initialize_cardtrak_summary()
  Dim i
  
  For i = 0 To MAX_CARDS
    cardtrak_summary(i).active = False
    cardtrak_summary(i).name = ""
    cardtrak_summary(i).balance = 0
    cardtrak_summary(i).interest = 0
    cardtrak_summary(i).late = 0
    cardtrak_summary(i).minimum = 0
    cardtrak_summary(i).paid = 0
    cardtrak_summary(i).purchases = 0
  Next i
  
  For i = 0 To 12
    cardtrak_monthly_summary(i).balance = 0
    cardtrak_monthly_summary(i).interest = 0
    cardtrak_monthly_summary(i).late = 0
    cardtrak_monthly_summary(i).minimum = 0
    cardtrak_monthly_summary(i).paid = 0
    cardtrak_monthly_summary(i).purchases = 0
  Next i

  cardtrak_summary_single_month.balance = 0
  cardtrak_summary_single_month.purchases = 0
  cardtrak_summary_single_month.interest = 0
  cardtrak_summary_single_month.late = 0
  cardtrak_summary_single_month.paid = 0
  cardtrak_summary_single_month.minimum = 0

End Sub


Public Sub put_this_in_cardtrak_summary(c As Integer)
  ' Total up the cardtrak records
  ' c=0 then summary is the sum of all cards for the month
  ' c>0 then only total this particular card
  
  Dim start As Integer
  Dim ending As Integer
  Dim n As Integer
  Dim qi As Integer
  Dim ct As card_transaction_type  ' This is where we'll put the card transaction
  Static last_qi As Integer
  
  qi = quick_index
  
  ' See if the current data matches where we are
  If (quick_index >= 0) And (quick_index <= 11) And (this.exclude = False) And (this.paid = PAID_DONE) Then
' esk I added the following line to show cardtracks as if they were paid. Messed things up so I commentedit out 2/2/2019
'esk  If (quick_index >= 0) And (quick_index <= 11) And (this.exclude = False) Then 'And (this.paid = PAID_DONE) Then
    ' We have a tab that it this record can go on
    If (is_cardtrak_transaction(this.name) And (this.sub_transaction_number > 0)) Then
      ' We have a ct record and we are on a month we want to keep
      n = Val(Mid(this.name, 4, 2))  ' Get the cardtrak number
      ct = cards(this.sub_transaction_number)
      
      ' Collect the monthly summary
      If (c = 0) Or (n = c) Then
        cardtrak_monthly_summary(qi).balance = cardtrak_monthly_summary(qi).balance + ct.new_balance
        cardtrak_monthly_summary(qi).purchases = cardtrak_monthly_summary(qi).purchases + ct.total_purchases
        cardtrak_monthly_summary(qi).interest = cardtrak_monthly_summary(qi).interest + ct.total_interest
        cardtrak_monthly_summary(qi).paid = cardtrak_monthly_summary(qi).paid + this.amount  'ct.amount_paid  '.total_payments
        cardtrak_monthly_summary(qi).late = cardtrak_monthly_summary(qi).late + ct.total_late
        cardtrak_monthly_summary(qi).minimum = cardtrak_monthly_summary(qi).minimum + ct.amount_due
      End If
      
      ' Collect the card summary if we are on the selected month
      If (this.Year = view.current_year) And (this.Month = view.current_month) Then
        ' We have a ct trans for the current active month
        cardtrak_summary(n).active = True
        cardtrak_summary(n).name = ct.name
        cardtrak_summary(n).balance = cardtrak_summary(n).balance + ct.new_balance
        cardtrak_summary(n).purchases = cardtrak_summary(n).purchases + ct.total_purchases
        cardtrak_summary(n).interest = cardtrak_summary(n).interest + ct.total_interest
        cardtrak_summary(n).late = cardtrak_summary(n).late + ct.total_late
        cardtrak_summary(n).paid = cardtrak_summary(n).paid + this.amount  ' ct.amount_paid  '.total_payments
        cardtrak_summary(n).minimum = cardtrak_summary(n).minimum + ct.amount_due
        
        cardtrak_summary_single_month.balance = cardtrak_summary_single_month.balance + cardtrak_summary(n).balance
        cardtrak_summary_single_month.purchases = cardtrak_summary_single_month.purchases + cardtrak_summary(n).purchases
        cardtrak_summary_single_month.interest = cardtrak_summary_single_month.interest + cardtrak_summary(n).interest
        cardtrak_summary_single_month.late = cardtrak_summary_single_month.late + cardtrak_summary(n).late
        cardtrak_summary_single_month.paid = cardtrak_summary_single_month.paid + cardtrak_summary(n).paid
        cardtrak_summary_single_month.minimum = cardtrak_summary_single_month.minimum + cardtrak_summary(n).minimum
      End If
    End If
    
    
  End If
  
  
End Sub

Public Function strip_off_ct_number(s As String) As String
  strip_off_ct_number = Mid(s, 8, 255)
End Function


