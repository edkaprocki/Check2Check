Attribute VB_Name = "Module6"
Option Explicit

' Define the words used in the program
Global Const FILE_N = 0
Global Const EDIT_N = 1
Global Const BEGINNING_N = 2
Global Const ENDING_N = 3
Global Const BALANCE_N = 4
Global Const INSERT_LAST_NAME_N = 5
Global Const LAST_REFERENCE_NUMBER_N = 6
Global Const NEW_REFERENCE_NUMBER_N = 7
Global Const CHECKBOOK_N = 8
Global Const NEW_N = 9
Global Const OPEN_N = 10
Global Const SAVE_N = 11
Global Const SAVE_AS_N = 12
Global Const VIEW_N = 13
Global Const PRINT_N = 14
Global Const PREFERENCES_N = 15
Global Const EXIT_N = 16
'Global Const EDIT_N = 17
Global Const UNDO_N = 18
Global Const EDIT_TRANSACTION_N = 19
Global Const MARK_TRANSACTION_WITH_N = 20
Global Const BLANK_N = 21
Global Const DONE_N = 22
Global Const PENDING_N = 23
Global Const SKIP_N = 24
Global Const TAGS_N = 25
Global Const SET_N = 26
Global Const ONE_N = 27
Global Const TWO_N = 28
Global Const THREE_N = 29
Global Const FOUR_N = 30
Global Const CLEAR_N = 31
'Global Const ONE_N = 32
'Global Const TWO_N = 33
'Global Const THREE_N = 34
'Global Const FOUR_N = 35
Global Const FILTER_TRANSACTIONS_N = 36
Global Const QUICK_SAVE_N = 37
Global Const QUICK_DEPOSIT_N = 38
Global Const QUICK_VIEW_EDIT_N = 39
Global Const INSERT_N = 40
Global Const DELETE_N = 41
Global Const COPY_N = 42
Global Const CUT_N = 43
Global Const PASTE_PENDING_N = 44
Global Const PASTE_AND_CLEAR_N = 45
Global Const PASTE_INTACT_N = 46
Global Const COPY_SELECTED_N = 47
Global Const PASTE_SELECTED_N = 48
Global Const PASTE_SELECTED_INTO_CURRENT_DATE_WITH_STATUS_N = 49
'Global Const BLANK_N = 50
Global Const INTACT_N = 51
'Global Const PENDING_N = 52
Global Const COPY_MONTH_N = 53
Global Const CUT_MONTH_N = 54
Global Const PASTE_MONTH_N = 55
Global Const PASTE_MONTH_AND_ARRANGE_N = 56
Global Const PASTE_MONTH_WITH_STATUS_N = 57
'Global Const BLANK_N = 58
'Global Const INTACT_N = 59
'Global Const PENDING_N = 60
Global Const COPY_TAGS_N = 61
Global Const CUT_TAGS_N = 62
Global Const PASTE_TAGS_AND_ARRANGE_WITH_STATUS_N = 63
Global Const PASTE_TAGS_INTO_CURRENT_DATE_WITH_STATUS_N = 64
Global Const PASTE_TAGS_WITH_STATUS_N = 65
Global Const INSERT_REFERENCE_NUMBER_N = 66
'Global Const CHECKBOOK_N = 67
Global Const RECONCILE_N = 68
Global Const CARDTRAK_N = 69
Global Const NEW_TRANSACTION_N = 70
'Global Const EDIT_TRANSACTION_N = 71
Global Const CONVERT_TO_CARDTRAK_N = 72
Global Const ADD_DELETE_EDIT_CARDS_N = 73
'Global Const VIEW_N = 74
Global Const QUICK_ACCOUNTS_N = 75
Global Const NEXT_MONTH_N = 76
Global Const PREVIOUS_MONTH_N = 77
Global Const TRANSACTIONS_N = 78
Global Const MONTHLY_NOTES_N = 79
Global Const CALENDAR_N = 80
Global Const CALCULATOR_N = 81
Global Const OVERRIDE_COLUMNS_N = 82
Global Const GO_TO_MONTH_YEAR_N = 83
Global Const BALANCES_N = 84
'Global Const TAGS_N = 85
Global Const SUMMARY_N = 86
Global Const CARDTRAK_SUMMARY_N = 87
Global Const HELP_N = 88
Global Const CONTENTS_N = 89
Global Const INDEX_N = 90
Global Const REGISTER_CHECK2CHECK_N = 91
Global Const BUY_NOW_N = 92
Global Const CHECK_FOR_LATEST_VERSION_VIA_INTERNET_N = 93
Global Const CHECK2CHECK_WEB_SITE = 94
Global Const LANGUAGE_N = 95
Global Const ENGLISH_N = 96
Global Const SPANISH_N = 97
Global Const ABOUT_N = 98
Global Const CLOSE_N = 99
Global Const CUT_SELECTED_N = 100
Global Const DATE_N = 101
Global Const DAY_N = 102
Global Const DUE_N = 103
Global Const CHECK_N = 104
Global Const NAME_N = 105
Global Const STATUS_N = 106
Global Const AMOUNT_N = 107
Global Const EXCLUDE_N = 108
Global Const CLR_N = 109
Global Const OR_N = 110
Global Const OR_BALANCE_N = 111
Global Const JAN_N = 112
Global Const FEB_N = 113
Global Const MAR_N = 114
Global Const APR_N = 115
Global Const MAY_N = 116
Global Const JUN_N = 117
Global Const JUL_N = 118
Global Const AUG_N = 119
Global Const SEP_N = 120
Global Const OCT_N = 121
Global Const NOV_N = 122
Global Const DEC_N = 123
Global Const JAN_LONG_N = 124
Global Const FEB_LONG_N = 125
Global Const MAR_LONG_N = 126
Global Const APR_LONG_N = 127
Global Const MAY_LONG_N = 128
Global Const JUN_LONG_N = 129
Global Const JUL_LONG_N = 130
Global Const AUG_LONG_N = 131
Global Const SEP_LONG_N = 132
Global Const OCT_LONG_N = 133
Global Const NOV_LONG_N = 134
Global Const DEC_LONG_N = 135
Global Const SUN_N = 136
Global Const MON_N = 137
Global Const TUE_N = 138
Global Const WED_N = 139
Global Const THU_N = 140
Global Const FRI_N = 141
Global Const SAT_N = 142
Global Const UNDO_CUT_SELECTED_N = 143
Global Const UNDO_CUT_N = 144
Global Const UNDO_EDIT_TRANSACTION_N = 145
Global Const UNDO_CUT_TRANSACTION_N = 146
Global Const UNDO_DELETE_TRANSACTION_N = 147
Global Const FILE_NOT_FOUND_N = 148
Global Const FILE_LENGTH_IS_ZERO_N = 149
Global Const UNDO_PASTE_TRANSACTION_N = 150
Global Const UNDO_MOVE_TRANSACTION_N = 151
Global Const UNDO_PASTE_N = 152
Global Const UNDO_PASTE_AND_CLEAR_N = 153
Global Const UNDO_MOVE_SELECTED_N = 154
Global Const WITHDRAWAL_N = 155
Global Const DEPOSIT_N = 156
Global Const CANCEL_N = 157
Global Const OK_N = 158
Global Const CLEAR_FILTER_N = 159
Global Const FILTER_N = 160
Global Const RECONCILE_CHECKBOOK_N = 161
Global Const BANK_STATEMENT_N = 162
Global Const BEGINNING_BALANCE_N = 163
Global Const ENDING_BALANCE_N = 164
Global Const CLEARED_BALANCE_N = 165
Global Const DIFFERENCE_N = 166
Global Const DEPOSITS_N = 167
Global Const CHECKS_N = 168
Global Const WITHDRAWALS_N = 169
Global Const SHOW_ALL_N = 170
Global Const NUMBER_N = 171
Global Const FINISH_N = 172
Global Const FINISH_LATER_N = 173
Global Const CLEARED_TRANSACTIONS_N = 174
Global Const ARE_YOU_SURE_N = 175
Global Const QUICK_ACCOUNTS_VIEW_EDIT_N = 176
Global Const DUE_DATE_N = 177
Global Const AMOUNT_NEEDED_N = 178
Global Const TOTAL_SAVED_N = 179
Global Const AMOUNT_SAVED_N = 180
Global Const PENDING_AMOUNT_N = 181
Global Const CREDIT_CARD_N = 182
Global Const PAYMENT_N = 183
Global Const ADD_NEW_CARD_N = 184
Global Const CARD_NAME_N = 185
Global Const NAME_ADDRESS_PHONE_N = 186
Global Const FINANCE_INFORMATION_N = 187
Global Const PREVIOUS_BALANCE_N = 188
Global Const PURCHASES_ADVANCES_N = 189
Global Const PAYMENTS_CREDITS_N = 190
Global Const FINANCE_CHARGES_N = 191
Global Const LATE_CHARGES_N = 192
Global Const CURRENT_BALANCE_N = 193
Global Const PAYMENT_INFORMATION_N = 194
Global Const CARD_DATE_DUE_N = 195
Global Const AMOUNT_DUE_N = 196
Global Const DATE_PAID_N = 197
Global Const AMOUNT_PAID_N = 198
Global Const CHECK_NUMBER_N = 199
Global Const POSTED_N = 200
Global Const INTEREST_N = 201
Global Const RATE_PERCENT_N = 202
Global Const CHARGES_N = 203
Global Const CLEAR_TRANSACTIONS_N = 204
Global Const CLEAR_INTEREST_N = 205
Global Const CLEAR_ALL_N = 206
Global Const DELETE_CARD_N = 207
Global Const DELETE_ALL_CARDS_N = 208
'Global Const CARD_SAVE_N = 209
Global Const SAVE_CHANGES_N = 210
Global Const YES_N = 211
Global Const NO_N = 212
Global Const LOW_N = 213
Global Const HIGH_N = 214
Global Const AVERAGE_N = 215
Global Const CARD_PURCHASES_N = 216
Global Const PAID_N = 217
Global Const LATE_N = 218
Global Const MINIMUM_N = 219
Global Const EST_BALANCE_N = 220
Global Const HIGH_LOW_N = 221
Global Const BEGIN_END_N = 222
Global Const DETAILED_SUMMARY_N = 223
Global Const REGISTER_N = 224
Global Const PURCHASE_INFORMATION_N = 225
Global Const RUN_CHECK2CHECK_N = 226
Global Const CODE_N = 227
Global Const NOT_ENOUGH_INFORMATION_N = 228
Global Const TOTAL_N = 229
Global Const QTY_N = 230
Global Const TO_N = 231
Global Const QTY_INC_N = 232
Global Const QTY_EXP_N = 233
Global Const INCOME_N = 234
Global Const EXPENSE_N = 235
Global Const REFERENCE_NUMBER_N = 236
Global Const NOTES_N = 237
Global Const CHARACTERS_LEFT_N = 238
Global Const NEXT_MONTH_TOP_N = 239
Global Const PREVIOUS_MONTH_BOTTOM_N = 240
Global Const EST_BAL_BAL_N = 241
Global Const CHANGE_N = 242
Global Const ALL_CARDS_N = 243
Global Const EST_BAL_N = 244
Global Const PURCHASES_N = 245
Global Const CLEARED_N = 246
Global Const FINISHED_N = 247
Global Const SORT_BY_N = 248
Global Const PROMPT_FOR_MOVE_N = 249
Global Const PROMPT_FOR_DELETE_N = 250
Global Const SHOW_NAMES_IN_COLORS_N = 251
Global Const AUTO_CHECK_DONE_ON_AMOUNT_N = 252
Global Const AUTO_INSERT_NEW_LINE_N = 253
Global Const SHOW_OVERRIDE_COLUMNS_N = 254
Global Const SHOW_SPLASH_SCREEN_N = 255
Global Const PROMPT_FOR_PASTE_NOTES_N = 256
Global Const AUTO_NEGATIVE_NUMBERS_N = 257
Global Const AUTO_CHECK_DONE_ON_CHECK_N = 258
Global Const SAVE_RECOVERY_FILE_N = 259
Global Const SAME_N = 260
Global Const CHANGE_PASSWORD_N = 261
Global Const FONT_SIZE_NOTES_N = 262
Global Const GENERAL_N = 263
Global Const STATUS_COLUMN_N = 264
Global Const CHECK_NUMBER_COLUMN_N = 265
Global Const CLEARED_COLUMN_N = 266
Global Const PASSWORD_N = 267
Global Const ENTER_NEW_PASSWORD_N = 268
Global Const ENTER_PASSWORD_N = 269
Global Const AUTO_LOAD_LAST_FILE_N = 270
Global Const NO_TRANSACTIONS_COPIED_N = 271
Global Const NO_TRANSACTIONS_TO_CUT_N = 272
Global Const CUT_ALL_SELECTED_Q_N = 273
Global Const TRANSACTIONS_FOR_N = 274
Global Const NO_TRANSACTIONS_TO_COPY_N = 275
Global Const CUT_ALL_TAGGED_Q_N = 276
Global Const DELETE_TRANSACTION_Q_N = 277
Global Const CUT_ALL_TRANSACTIONS_FOR_N = 278
Global Const MOVE_TRANSACTION_Q_N = 279
Global Const CANT_ASSIGN_A_CHECK_NUMBER_TO_A_DEPOSIT_N = 280
Global Const CHECK_NUMBER_ERROR_N = 281
Global Const SAVE_DATABASE_Q_N = 282
Global Const ERROR_N = 283
Global Const PASTE_ALL_TRANSACTIONS_TO_N = 284
Global Const PASTE_NOTES_Q_N = 285
Global Const ADJUSTED_DATES_TO_MATCH_THE_CURRENT_MONTH_N = 286
Global Const CANT_PASTE_TO_THE_SAME_MONTH_AND_YEAR_N = 287
Global Const NO_TRANSACTIONS_CUT_N = 288
Global Const TRANSACTION_FOR_N = 289
Global Const FILE_EXISTS_OVERWRITE_Q_N = 290
Global Const OVERWRITE_Q_N = 291
Global Const ERROR_IN_DELETING_WRITING_FILE_N = 292
Global Const INVALID_NUMBER_ENTERED_N = 293
Global Const INVALID_TAG_NUMBER_N = 294
Global Const CHANGING_THIS_MAY_CAUSE_A_PROBLEM_WHEN_YOU_RECONCILE_Q_N = 295
Global Const INVALID_KEY_N = 296
Global Const WARNING_QUICK_SAVE_ACCOUNT_BALANCE_IS_NOT_ZERO_N = 297
Global Const CALCULATOR_NOT_FOUND_N = 298
Global Const SAVE_CHANGES_Q_N = 299
Global Const DELETE_ALL_CREDIT_CARDS_AND_INFORMATION_Q_N = 300
Global Const ARE_YOU_ABSOLUTELY_SURE_Q_N = 301
Global Const ALL_CREDIT_CARDS_HAVE_BEEN_DELETED_N = 302
Global Const DELETE_Q_N = 303
Global Const NEW_TRANSACTION_ENTERED_N = 304
Global Const MUST_HAVE_A_TRANSACTION_NAME_N = 305
Global Const MUST_SELECT_A_CREDIT_CARD_N = 306
Global Const SELECT_CARD_N = 307
Global Const ERRORS_IN_CT_DATABASE_N = 308
Global Const ADD_NEW_CARD_Q_N = 309
Global Const SORRY_NO_SPACE_FOR_CARDS_N = 310
Global Const ARE_YOU_SURE_Q_N = 311
Global Const INVALID_ENTRY_N = 312
Global Const PASSWORDS_DONT_MATCH_N = 313
Global Const DELETE_QUICK_SAVE_ACCOUNT_Q_N = 314
Global Const ARE_YOU_SURE_Q_QUICK_SAVE_ACCOUNT_IS_NOT_ZERO_N = 315
Global Const INVALID_NUMBER_N = 316
Global Const SORRY_NOT_ENOUGH_MONEY_IN_THAT_ACCOUNT_N = 317
Global Const SAVE_QUICK_ACCOUNT_DATA_Q_N = 318
Global Const INVALID_AMOUNT_N = 319
Global Const DO_YOU_WANT_TO_FINALIZE_Q_N = 320
Global Const THE_DIFFERENCE_IS_NOT_ZERO_Q_N = 321
Global Const SUCCESS_THANK_YOU_FOR_REGISTERING_N = 322
Global Const SORRY_NOT_A_VALID_NAME_CODE_N = 323
Global Const DELETE_QUICK_SAVE_ACCOUNT_N = 324
Global Const CAUTION_TRANSACTION_PERMANENTLY_MARKED_N = 325
Global Const SAVE_Q_N = 326
Global Const ADD_Q_N = 327
Global Const DELETE_ACCOUNT_Q_N = 328
Global Const SELECT_THIS_CARD_Q_N = 329
Global Const CHANGE_FINISHED_STATUS_Q_N = 330
Global Const SAVE_QUICK_ACCOUNT_DATA_N = 331
Global Const CANCEL_RECONCILIATION_Q_N = 332
Global Const SUCCESS_RECONCILIATION_COMPLETE_N = 333
Global Const CAUTION_RECONCILIATION_NOT_COMPLETE_N = 334


Global Const CREDIT_CARD_TRANSACTIONS_N = 335
Global Const CARD_N = 336
Global Const DELETE_THIS_CARD_N = 337
Global Const DELETE_ALL_CARDS_MENU_N = 338
Global Const SELECT_THIS_CARD_N = 339
Global Const SELECT_N = 340
Global Const NAME_ADDRESS_PHONE_NUMBERS_N = 341
Global Const DATE_DUE_N = 342
Global Const CREDIT_CARD_INFORMATION_N = 343
Global Const FIRST_TRANSACTION_N = 344
Global Const LAST_TRANSACTION_N = 345
Global Const TOTAL_TRANSACTIONS_N = 346
Global Const NEXT_CHECK_NUMBER_N = 347
Global Const VERSION_N = 348
Global Const ACCOUNT_N = 349
Global Const TOTAL_NOTES_N = 350
Global Const TOTAL_CARDTRAKS_N = 351
Global Const PLAY_SOUNDS_N = 352
Global Const PASTE_N = 353
Global Const PASTE_OPTIONS_N = 354
Global Const PASTE_TAGS_OPTIONS_N = 355
Global Const MONTH_N = 356
Global Const TOTAL_DELTA_N = 357
Global Const DELTA_N = 358

Dim word_list(358, 1)     ' Make this number the same as the line above

Dim words_are_initialized As Boolean  ' True when table has been set up
 
 
 
Public Function words(n As Integer) As String
  If (Not words_are_initialized) Then words_initialize
  words = word_list(n, preferences.language)
  
  words_are_initialized = True
End Function


 
Public Sub words_initialize()
  word_list(FILE_N, 0) = "File"
  word_list(FILE_N, 1) = "Archivo"

  word_list(EDIT_N, 0) = "Edit"
  word_list(EDIT_N, 1) = "Editar"

  word_list(BEGINNING_N, 0) = "Beginning"
  word_list(BEGINNING_N, 1) = "Principio"

  word_list(ENDING_N, 0) = "Ending"
  word_list(ENDING_N, 1) = "Final"

  word_list(BALANCE_N, 0) = "Balance"
  word_list(BALANCE_N, 1) = "Balance"

  word_list(INSERT_LAST_NAME_N, 0) = "Insert Last Name"
  word_list(INSERT_LAST_NAME_N, 1) = "Nombre Actual"

  word_list(LAST_REFERENCE_NUMBER_N, 0) = "Last Ref Number"
  word_list(LAST_REFERENCE_NUMBER_N, 1) = "Número Actual"

  word_list(NEW_REFERENCE_NUMBER_N, 0) = "New Ref Number"
  word_list(NEW_REFERENCE_NUMBER_N, 1) = "Número Nuevo"

  word_list(CHECKBOOK_N, 0) = "Checkbook"
  word_list(CHECKBOOK_N, 1) = "Chequera"

  word_list(NEW_N, 0) = "New"
  word_list(NEW_N, 1) = "Nuevo"

  word_list(OPEN_N, 0) = "Open"
  word_list(OPEN_N, 1) = "Abrir"

  word_list(SAVE_N, 0) = "Save"
  word_list(SAVE_N, 1) = "Guardar"

  word_list(SAVE_AS_N, 0) = "Save As"
  word_list(SAVE_AS_N, 1) = "Guardar como"

  word_list(VIEW_N, 0) = "View"
  word_list(VIEW_N, 1) = "Ver"

  word_list(PRINT_N, 0) = "Print"
  word_list(PRINT_N, 1) = "Imprimir"

  word_list(PREFERENCES_N, 0) = "Preferences"
  word_list(PREFERENCES_N, 1) = "Preferencias"

  word_list(EXIT_N, 0) = "Exit"
  word_list(EXIT_N, 1) = "Salir"

  word_list(UNDO_N, 0) = "Undo"
  word_list(UNDO_N, 1) = "Deshacer"

  word_list(EDIT_TRANSACTION_N, 0) = "Edit Transaction"
  word_list(EDIT_TRANSACTION_N, 1) = "Editar transacción"

  word_list(MARK_TRANSACTION_WITH_N, 0) = "Mark transaction with..."
  word_list(MARK_TRANSACTION_WITH_N, 1) = "Señalar transacción con..."

  word_list(BLANK_N, 0) = "Blank"
  word_list(BLANK_N, 1) = "en Blanco"

  word_list(DONE_N, 0) = "Done"
  word_list(DONE_N, 1) = "Terminado"

  word_list(PENDING_N, 0) = "Pending"
  word_list(PENDING_N, 1) = "Pendiente"

  word_list(SKIP_N, 0) = "Skip"
  word_list(SKIP_N, 1) = "Brincar"

  word_list(TAGS_N, 0) = "Tags"
  word_list(TAGS_N, 1) = "Etiquetas"

  word_list(SET_N, 0) = "Set"
  word_list(SET_N, 1) = "poner"

  word_list(ONE_N, 0) = "1"
  word_list(ONE_N, 1) = "1"

  word_list(TWO_N, 0) = "2"
  word_list(TWO_N, 1) = "2"

  word_list(THREE_N, 0) = "3"
  word_list(THREE_N, 1) = "3"

  word_list(FOUR_N, 0) = "4"
  word_list(FOUR_N, 1) = "4"

  word_list(CLEAR_N, 0) = "Clear"
  word_list(CLEAR_N, 1) = "Vaciar"

  word_list(FILTER_TRANSACTIONS_N, 0) = "Filter transactions"
  word_list(FILTER_TRANSACTIONS_N, 1) = "Filtro"

  word_list(QUICK_SAVE_N, 0) = "Quick Save"
  word_list(QUICK_SAVE_N, 1) = "Guardado Rápido"

  word_list(QUICK_DEPOSIT_N, 0) = "Quick Deposit"
  word_list(QUICK_DEPOSIT_N, 1) = "Deposito rápido"

  word_list(QUICK_VIEW_EDIT_N, 0) = "Quick View/Edit"
  word_list(QUICK_VIEW_EDIT_N, 1) = "Rápido"

  word_list(INSERT_N, 0) = "Insert"
  word_list(INSERT_N, 1) = "Insertar"

  word_list(DELETE_N, 0) = "Delete"
  word_list(DELETE_N, 1) = "Borrar"

  word_list(COPY_N, 0) = "Copy"
  word_list(COPY_N, 1) = "Copiar"

  word_list(CUT_N, 0) = "Cut"
  word_list(CUT_N, 1) = "Cortar"

  word_list(PASTE_PENDING_N, 0) = "Paste Pending"
  word_list(PASTE_PENDING_N, 1) = "Pegado pendiente"

  word_list(PASTE_AND_CLEAR_N, 0) = "Paste and Clear"
  word_list(PASTE_AND_CLEAR_N, 1) = "Pegar y Vaciar"

  word_list(PASTE_INTACT_N, 0) = "Paste Intact"
  word_list(PASTE_INTACT_N, 1) = "Pegar Intacto"

  word_list(COPY_SELECTED_N, 0) = "Copy Selected"
  word_list(COPY_SELECTED_N, 1) = "Copiar Selección"

  word_list(PASTE_SELECTED_N, 0) = "Paste Selected"
  word_list(PASTE_SELECTED_N, 1) = "Pegar Selección"

  
  word_list(PASTE_SELECTED_INTO_CURRENT_DATE_WITH_STATUS_N, 0) = "Paste Selected into current date with Status..."
  word_list(PASTE_SELECTED_INTO_CURRENT_DATE_WITH_STATUS_N, 1) = "Pegado Seleccionado en el día actual con Condición..."

  word_list(INTACT_N, 0) = "Intact"
  word_list(INTACT_N, 1) = "Intacto"

  word_list(COPY_MONTH_N, 0) = "Copy Month"
  word_list(COPY_MONTH_N, 1) = "Copiar Mes"

  word_list(CUT_MONTH_N, 0) = "Cut Month"
  word_list(CUT_MONTH_N, 1) = "Cortar Mes"

  word_list(PASTE_MONTH_N, 0) = "Paste Month"
  word_list(PASTE_MONTH_N, 1) = "Pegar Mes"

  word_list(PASTE_MONTH_AND_ARRANGE_N, 0) = "Paste Month and Arrange"
  word_list(PASTE_MONTH_AND_ARRANGE_N, 1) = "Pegar mes y arreglar"

  word_list(PASTE_MONTH_WITH_STATUS_N, 0) = "Paste Month with Status..."
  word_list(PASTE_MONTH_WITH_STATUS_N, 1) = "Pegar mes con Condición..."

  word_list(COPY_TAGS_N, 0) = "Copy Tags"
  word_list(COPY_TAGS_N, 1) = "Copiar Etiquetas"

  word_list(CUT_TAGS_N, 0) = "Cut Tags"
  word_list(CUT_TAGS_N, 1) = "Cortar Etiquetas"

  
  word_list(PASTE_TAGS_AND_ARRANGE_WITH_STATUS_N, 0) = "Paste Tags and Arrange with Status..."
  word_list(PASTE_TAGS_AND_ARRANGE_WITH_STATUS_N, 1) = "Pegar Etiquetas y Arreglar con Condición..."
  
  word_list(PASTE_TAGS_INTO_CURRENT_DATE_WITH_STATUS_N, 0) = "Paste Tags into Current Date with Status..."
  word_list(PASTE_TAGS_INTO_CURRENT_DATE_WITH_STATUS_N, 1) = "Pegar Etiquetas en el día Actual con Condición..."

  word_list(PASTE_TAGS_WITH_STATUS_N, 0) = "Paste Tags with Status..."
  word_list(PASTE_TAGS_WITH_STATUS_N, 1) = "Pegar Etiquetas con Condición..."

  word_list(INSERT_REFERENCE_NUMBER_N, 0) = "Insert Reference Number"
  word_list(INSERT_REFERENCE_NUMBER_N, 1) = "Número de Referencia"

  word_list(REFERENCE_NUMBER_N, 0) = "Reference Number"
  word_list(REFERENCE_NUMBER_N, 1) = "Número de Referencia"

  word_list(RECONCILE_N, 0) = "Reconcile"
  word_list(RECONCILE_N, 1) = "Reconciliar"

  word_list(CARDTRAK_N, 0) = "Cardtrak"
  word_list(CARDTRAK_N, 1) = "Rastreo-Tarjeta"

  word_list(NEW_TRANSACTION_N, 0) = "New Transaction"
  word_list(NEW_TRANSACTION_N, 1) = "Nueva Transacción"

  word_list(EDIT_TRANSACTION_N, 0) = "Edit Transaction"
  word_list(EDIT_TRANSACTION_N, 1) = "Editar Transacción"

  word_list(CONVERT_TO_CARDTRAK_N, 0) = "Convert to Cardtrak"
  word_list(CONVERT_TO_CARDTRAK_N, 1) = "Convertir a Rastreo-Tarjeta"

  word_list(ADD_DELETE_EDIT_CARDS_N, 0) = "Add/Delete/Edit Cards"
  word_list(ADD_DELETE_EDIT_CARDS_N, 1) = "Agregar/Borrar/Editar Tarjetas"

  word_list(QUICK_ACCOUNTS_N, 0) = "Quick Accounts"
  word_list(QUICK_ACCOUNTS_N, 1) = "Cuenta Rápida"

  word_list(NEXT_MONTH_N, 0) = "Next Month"
  word_list(NEXT_MONTH_N, 1) = "Mes Prox"

  word_list(PREVIOUS_MONTH_N, 0) = "Previous Month"
  word_list(PREVIOUS_MONTH_N, 1) = "Mes Anterior"

  word_list(TRANSACTIONS_N, 0) = "Transactions"
  word_list(TRANSACTIONS_N, 1) = "Transacciones"

  word_list(MONTHLY_NOTES_N, 0) = "Monthly Notes"
  word_list(MONTHLY_NOTES_N, 1) = "Notas Mensuales"

  word_list(NOTES_N, 0) = "Notes"
  word_list(NOTES_N, 1) = "Notas"

  word_list(CALENDAR_N, 0) = "Calendar"
  word_list(CALENDAR_N, 1) = "Calendario"

  word_list(CALCULATOR_N, 0) = "Calculator"
  word_list(CALCULATOR_N, 1) = "Calculadora"

  word_list(OVERRIDE_COLUMNS_N, 0) = "Override Columns"
  word_list(OVERRIDE_COLUMNS_N, 1) = "Ignorar columnas"

  word_list(GO_TO_MONTH_YEAR_N, 0) = "Go to Month / Year"
  word_list(GO_TO_MONTH_YEAR_N, 1) = "Ir a Mes / Año"

  word_list(BALANCES_N, 0) = "Balances"
  word_list(BALANCES_N, 1) = "Balances"

  word_list(SUMMARY_N, 0) = "Summary"
  word_list(SUMMARY_N, 1) = "Sumario"

  word_list(CARDTRAK_SUMMARY_N, 0) = "Cardtrak Summary"
  word_list(CARDTRAK_SUMMARY_N, 1) = "Sumario del Rastreo de Tarjeta"

  word_list(HELP_N, 0) = "Help"
  word_list(HELP_N, 1) = "Ayuda"

  word_list(CONTENTS_N, 0) = "Contents"
  word_list(CONTENTS_N, 1) = "Contenido"

  word_list(INDEX_N, 0) = "Index"
  word_list(INDEX_N, 1) = "Índice"

  word_list(REGISTER_CHECK2CHECK_N, 0) = "Register Check2Check"
  word_list(REGISTER_CHECK2CHECK_N, 1) = "Registrar Check2Check"

  word_list(BUY_NOW_N, 0) = "Buy Now"
  word_list(BUY_NOW_N, 1) = "Comprar Ahora"

  
  word_list(CHECK_FOR_LATEST_VERSION_VIA_INTERNET_N, 0) = "Check for Updates"
  word_list(CHECK_FOR_LATEST_VERSION_VIA_INTERNET_N, 1) = "Checar para última versión vía Internet"

  word_list(CHECK2CHECK_WEB_SITE, 0) = "Check2Check Web Site"
  word_list(CHECK2CHECK_WEB_SITE, 1) = "Sitio Web Check2Check"

  word_list(LANGUAGE_N, 0) = "Language / Idioma"
  word_list(LANGUAGE_N, 1) = "Language / Idioma"

  word_list(ENGLISH_N, 0) = "English"
  word_list(ENGLISH_N, 1) = "English"

  word_list(SPANISH_N, 0) = "Espanole"  'Spanish"
  word_list(SPANISH_N, 1) = "Espanole"

  word_list(ABOUT_N, 0) = "About"
  word_list(ABOUT_N, 1) = "Acerca de"

  word_list(CLOSE_N, 0) = "Close"
  word_list(CLOSE_N, 1) = "Cerrar"

  word_list(CUT_SELECTED_N, 0) = "Cut Selected"
  word_list(CUT_SELECTED_N, 1) = "Cortar Selección"

  word_list(DATE_N, 0) = "Date"
  word_list(DATE_N, 1) = "Fecha"

  word_list(DAY_N, 0) = "Day"
  word_list(DAY_N, 1) = "Día"

  word_list(DUE_N, 0) = "Due"
  word_list(DUE_N, 1) = "Debido"

  word_list(CHECK_N, 0) = "Check"
  word_list(CHECK_N, 1) = "Cheque"

  word_list(NAME_N, 0) = "Name"
  word_list(NAME_N, 1) = "Nombre"

  word_list(STATUS_N, 0) = "Status"
  word_list(STATUS_N, 1) = "Estado"

  word_list(AMOUNT_N, 0) = "Amount"
  word_list(AMOUNT_N, 1) = "Cantidad"

  word_list(EXCLUDE_N, 0) = "Excl"
  word_list(EXCLUDE_N, 1) = "Excluir"

  word_list(CLR_N, 0) = "CLR"
  word_list(CLR_N, 1) = "CLR"

  word_list(OR_N, 0) = "O/R"
  word_list(OR_N, 1) = "O/R"

  word_list(OR_BALANCE_N, 0) = "O/R Balance"
  word_list(OR_BALANCE_N, 1) = "O/R Balance"

  word_list(JAN_N, 0) = "Jan"
  word_list(JAN_N, 1) = "Ene"

  word_list(FEB_N, 0) = "Feb"
  word_list(FEB_N, 1) = "Feb"

  word_list(MAR_N, 0) = "Mar"
  word_list(MAR_N, 1) = "Mar"

  word_list(APR_N, 0) = "Apr"
  word_list(APR_N, 1) = "Abr"

  word_list(MAY_N, 0) = "May"
  word_list(MAY_N, 1) = "Mayo"

  word_list(JUN_N, 0) = "Jun"
  word_list(JUN_N, 1) = "Jun"

  word_list(JUL_N, 0) = "Jul"
  word_list(JUL_N, 1) = "Jul"

  word_list(AUG_N, 0) = "Aug"
  word_list(AUG_N, 1) = "Ago"

  word_list(SEP_N, 0) = "Sep"
  word_list(SEP_N, 1) = "Sep"

  word_list(OCT_N, 0) = "Oct"
  word_list(OCT_N, 1) = "Oct"

  word_list(NOV_N, 0) = "Nov"
  word_list(NOV_N, 1) = "Nov"

  word_list(DEC_N, 0) = "Dec"
  word_list(DEC_N, 1) = "Dic"

  word_list(JAN_LONG_N, 0) = "January"
  word_list(JAN_LONG_N, 1) = "Enero"

  word_list(FEB_LONG_N, 0) = "February"
  word_list(FEB_LONG_N, 1) = "Febrero"

  word_list(MAR_LONG_N, 0) = "March"
  word_list(MAR_LONG_N, 1) = "Marzo"

  word_list(APR_LONG_N, 0) = "April"
  word_list(APR_LONG_N, 1) = "Abril"

  word_list(MAY_LONG_N, 0) = "May"
  word_list(MAY_LONG_N, 1) = "Mayo"

  word_list(JUN_LONG_N, 0) = "June"
  word_list(JUN_LONG_N, 1) = "Junio"

  word_list(JUL_LONG_N, 0) = "July"
  word_list(JUL_LONG_N, 1) = "Julio"

  word_list(AUG_LONG_N, 0) = "August"
  word_list(AUG_LONG_N, 1) = "Agosto"

  word_list(SEP_LONG_N, 0) = "September"
  word_list(SEP_LONG_N, 1) = "Septiembre"

  word_list(OCT_LONG_N, 0) = "October"
  word_list(OCT_LONG_N, 1) = "Octubre"

  word_list(NOV_LONG_N, 0) = "November"
  word_list(NOV_LONG_N, 1) = "Noviembre"

  word_list(DEC_LONG_N, 0) = "December"
  word_list(DEC_LONG_N, 1) = "Diciembre"

  word_list(SUN_N, 0) = "Sun"
  word_list(SUN_N, 1) = "Dom"

  word_list(MON_N, 0) = "Mon"
  word_list(MON_N, 1) = "Lun"

  word_list(TUE_N, 0) = "Tue"
  word_list(TUE_N, 1) = "Mar"

  word_list(WED_N, 0) = "Wed"
  word_list(WED_N, 1) = "Mié"

  word_list(THU_N, 0) = "Thu"
  word_list(THU_N, 1) = "Jue"

  word_list(FRI_N, 0) = "Fri"
  word_list(FRI_N, 1) = "Vie"

  word_list(SAT_N, 0) = "Sat"
  word_list(SAT_N, 1) = "Sáb"

  word_list(UNDO_CUT_SELECTED_N, 0) = "Undo - Cut Selected"
  word_list(UNDO_CUT_SELECTED_N, 1) = "Deshacer - corte de selección"

  word_list(UNDO_CUT_N, 0) = "Undo - Cut"
  word_list(UNDO_CUT_N, 1) = "Deshacer - corte"

  word_list(UNDO_EDIT_TRANSACTION_N, 0) = "Undo Edit Transaction"
  word_list(UNDO_EDIT_TRANSACTION_N, 1) = "Deshacer transacción editada"

  word_list(UNDO_CUT_TRANSACTION_N, 0) = "Undo - Cut transaction"
  word_list(UNDO_CUT_TRANSACTION_N, 1) = "Deshacer - Corte de transacción"

  word_list(UNDO_DELETE_TRANSACTION_N, 0) = "Undo - Delete transaction"
  word_list(UNDO_DELETE_TRANSACTION_N, 1) = "Deshacer - Eliminación de"

  word_list(FILE_NOT_FOUND_N, 0) = "File not found"
  word_list(FILE_NOT_FOUND_N, 1) = "Archivo no encontrado"

  word_list(FILE_LENGTH_IS_ZERO_N, 0) = "File length is zero"
  word_list(FILE_LENGTH_IS_ZERO_N, 1) = "Archivo vació"

  word_list(UNDO_PASTE_TRANSACTION_N, 0) = "Undo - Paste transaction"
  word_list(UNDO_PASTE_TRANSACTION_N, 1) = "Deshacer - Pegar la transacción"

  word_list(UNDO_MOVE_TRANSACTION_N, 0) = "Undo - Move transaction"
  word_list(UNDO_MOVE_TRANSACTION_N, 1) = "Deshacer - Mover la transacción"

  word_list(UNDO_PASTE_N, 0) = "Undo - Paste"
  word_list(UNDO_PASTE_N, 1) = "Deshacer - pegar"

  word_list(UNDO_PASTE_AND_CLEAR_N, 0) = "Undo - Paste and clear"
  word_list(UNDO_PASTE_AND_CLEAR_N, 1) = "Deshacer  - pegar y Vaciar"

  word_list(UNDO_MOVE_SELECTED_N, 0) = "Undo - Move Selected"
  word_list(UNDO_MOVE_SELECTED_N, 1) = "Deshacer - Ultimo movimiento"

  word_list(WITHDRAWAL_N, 0) = "Withdrawal"
  word_list(WITHDRAWAL_N, 1) = "Retiro"

  word_list(DEPOSIT_N, 0) = "Deposit"
  word_list(DEPOSIT_N, 1) = "Depósito"

  word_list(CANCEL_N, 0) = "Cancel"
  word_list(CANCEL_N, 1) = "Cancelación"

  word_list(OK_N, 0) = "OK"
  word_list(OK_N, 1) = "OK"

  word_list(CLEAR_FILTER_N, 0) = "Clear Filter"
  word_list(CLEAR_FILTER_N, 1) = "Vaciar filtro"

  word_list(FILTER_N, 0) = "Filter"
  word_list(FILTER_N, 1) = "Filtro"

  word_list(RECONCILE_CHECKBOOK_N, 0) = "Reconcile Checkbook"
  word_list(RECONCILE_CHECKBOOK_N, 1) = "Reconciliación de Chequera"

  word_list(BANK_STATEMENT_N, 0) = "Bank Statement"
  word_list(BANK_STATEMENT_N, 1) = "Estado Bancario"

  word_list(BEGINNING_BALANCE_N, 0) = "Beginning Balance"
  word_list(BEGINNING_BALANCE_N, 1) = "Inicio de Balance"

  word_list(ENDING_BALANCE_N, 0) = "Ending Balance"
  word_list(ENDING_BALANCE_N, 1) = "Fin del Balance"

  word_list(CLEARED_BALANCE_N, 0) = "Cleared Balance"
  word_list(CLEARED_BALANCE_N, 1) = "Balance aclarado"

  word_list(DIFFERENCE_N, 0) = "Difference"
  word_list(DIFFERENCE_N, 1) = "Diferencia"

  word_list(DEPOSITS_N, 0) = "Deposits"
  word_list(DEPOSITS_N, 1) = "Depósitos"

  word_list(CHECKS_N, 0) = "Checks"
  word_list(CHECKS_N, 1) = "Cheques"

  word_list(WITHDRAWALS_N, 0) = "Withdrawals"
  word_list(WITHDRAWALS_N, 1) = "Retiros"

  word_list(SHOW_ALL_N, 0) = "Show All"
  word_list(SHOW_ALL_N, 1) = "Demuestra todo"

  word_list(NUMBER_N, 0) = "Number"
  word_list(NUMBER_N, 1) = "Nùmero"

  word_list(FINISH_N, 0) = "Finish"
  word_list(FINISH_N, 1) = "Terminar"

  word_list(FINISH_LATER_N, 0) = "Finish Later"
  word_list(FINISH_LATER_N, 1) = "Terminar después"

  word_list(CLEARED_TRANSACTIONS_N, 0) = "Cleared Transactions"
  word_list(CLEARED_TRANSACTIONS_N, 1) = "Transacciones aclaradas"

  word_list(ARE_YOU_SURE_N, 0) = "Are you sure?"
  word_list(ARE_YOU_SURE_N, 1) = "¿Seguro?"

  word_list(ARE_YOU_SURE_Q_N, 0) = "Are you sure?"
  word_list(ARE_YOU_SURE_Q_N, 1) = "¿Seguro?"

  word_list(QUICK_ACCOUNTS_VIEW_EDIT_N, 0) = "Quick Accounts - View/Edit"
  word_list(QUICK_ACCOUNTS_VIEW_EDIT_N, 1) = "Cuentas Rápidas - Ver/Editar"

  word_list(DUE_DATE_N, 0) = "Due Date"
  word_list(DUE_DATE_N, 1) = "Fecha de Adeudo"

  word_list(AMOUNT_NEEDED_N, 0) = "Amount Needed"
  word_list(AMOUNT_NEEDED_N, 1) = "Cantidad requerida"

  word_list(TOTAL_SAVED_N, 0) = "Total Saved"
  word_list(TOTAL_SAVED_N, 1) = "Total Ahorrado"

  word_list(AMOUNT_SAVED_N, 0) = "Amount Saved"
  word_list(AMOUNT_SAVED_N, 1) = "Cantidad Ahorrada"

  word_list(PENDING_AMOUNT_N, 0) = "Pending Amount"
  word_list(PENDING_AMOUNT_N, 1) = "Cantidad Pendiente"

  word_list(CREDIT_CARD_N, 0) = "Credit Card"
  word_list(CREDIT_CARD_N, 1) = "Tarjeta de Crédito"

  word_list(PAYMENT_N, 0) = "Payment"
  word_list(PAYMENT_N, 1) = "Pago"

  word_list(ADD_NEW_CARD_N, 0) = "Add New Card"
  word_list(ADD_NEW_CARD_N, 1) = "Añadir nueva tarjeta"

  word_list(CARD_NAME_N, 0) = "Card Name"
  word_list(CARD_NAME_N, 1) = "Nombre de la tarjeta"

  word_list(NAME_ADDRESS_PHONE_N, 0) = "Name/Address/Phone Numbers/Credit Limits/Comments"
  word_list(NAME_ADDRESS_PHONE_N, 1) = "Nombre/Dirección/Teléfono/Límites de Crédito/Comentarios"

  word_list(FINANCE_INFORMATION_N, 0) = "Finance Information"
  word_list(FINANCE_INFORMATION_N, 1) = "Datos Financieros"

  word_list(PREVIOUS_BALANCE_N, 0) = "Previous Balance"
  word_list(PREVIOUS_BALANCE_N, 1) = "Balance Previo"

  word_list(PURCHASES_ADVANCES_N, 0) = "Purchases & Advances"
  word_list(PURCHASES_ADVANCES_N, 1) = "Adquisiciones y Avances"

  word_list(PURCHASES_N, 0) = "Purchases"
  word_list(PURCHASES_N, 1) = "Adquisiciones"
  
  word_list(PAYMENTS_CREDITS_N, 0) = "Payments & Credits"
  word_list(PAYMENTS_CREDITS_N, 1) = "Pagos y Créditos"

  word_list(FINANCE_CHARGES_N, 0) = "Finance Charges"
  word_list(FINANCE_CHARGES_N, 1) = "Cargos Financieros"

  word_list(LATE_CHARGES_N, 0) = "Late Charges"
  word_list(LATE_CHARGES_N, 1) = "Penalizaciones"

  word_list(CURRENT_BALANCE_N, 0) = "Current Balance"
  word_list(CURRENT_BALANCE_N, 1) = "Balance Actual"

  word_list(PAYMENT_INFORMATION_N, 0) = "Payment Information"
  word_list(PAYMENT_INFORMATION_N, 1) = "Datos del Pago"

  word_list(CARD_DATE_DUE_N, 0) = "Date Due"
  word_list(CARD_DATE_DUE_N, 1) = "Fecha de Adeudo"

  word_list(AMOUNT_DUE_N, 0) = "Amount Due"
  word_list(AMOUNT_DUE_N, 1) = "Cantidad de Adeudo"

  word_list(DATE_PAID_N, 0) = "Date Paid"
  word_list(DATE_PAID_N, 1) = "Fecha de pago"

  word_list(AMOUNT_PAID_N, 0) = "Amount Paid"
  word_list(AMOUNT_PAID_N, 1) = "Cantidad del pago"

  word_list(CHECK_NUMBER_N, 0) = "Check Number"
  word_list(CHECK_NUMBER_N, 1) = "Nùmero de Cheque"

  word_list(POSTED_N, 0) = "Posted"
  word_list(POSTED_N, 1) = "Registrado"

  word_list(INTEREST_N, 0) = "Interest"
  word_list(INTEREST_N, 1) = "Interés"

  word_list(RATE_PERCENT_N, 0) = "Rate %"
  word_list(RATE_PERCENT_N, 1) = "Tarifa %"

  word_list(CHARGES_N, 0) = "Charges"
  word_list(CHARGES_N, 1) = "Cargos"

  word_list(CLEAR_TRANSACTIONS_N, 0) = "Clear Transactions"
  word_list(CLEAR_TRANSACTIONS_N, 1) = "Vaciar Transacciones"

  word_list(CLEAR_INTEREST_N, 0) = "Clear Interest"
  word_list(CLEAR_INTEREST_N, 1) = "Vaciar Intereses"

  word_list(CLEAR_ALL_N, 0) = "Clear All"
  word_list(CLEAR_ALL_N, 1) = "Vaciar Todo"

  word_list(DELETE_CARD_N, 0) = "Delete Card"
  word_list(DELETE_CARD_N, 1) = "Eliminar Tarjeta"

  word_list(DELETE_ALL_CARDS_N, 0) = "Delete All Cards"
  word_list(DELETE_ALL_CARDS_N, 1) = "Eliminar todas las tarjetas"

  word_list(SAVE_CHANGES_N, 0) = "Save Changes?"
  word_list(SAVE_CHANGES_N, 1) = "¿Guardar cambios?"

  word_list(YES_N, 0) = "Yes"
  word_list(YES_N, 1) = "Si"

  word_list(NO_N, 0) = "No"
  word_list(NO_N, 1) = "No"

  word_list(LOW_N, 0) = "Low"
  word_list(LOW_N, 1) = "Bajo"

  word_list(HIGH_N, 0) = "High"
  word_list(HIGH_N, 1) = "Alto"

  word_list(AVERAGE_N, 0) = "Average"
  word_list(AVERAGE_N, 1) = "Promedio"

  word_list(CARD_PURCHASES_N, 0) = "Purchases"
  word_list(CARD_PURCHASES_N, 1) = "Adquisiciones"

  word_list(PAID_N, 0) = "Paid"
  word_list(PAID_N, 1) = "Pagado"

  word_list(LATE_N, 0) = "Late"
  word_list(LATE_N, 1) = "Tardado"

  word_list(MONTH_N, 0) = "Month"
  word_list(MONTH_N, 1) = "Month"

  word_list(DELTA_N, 0) = "Delta"
  word_list(DELTA_N, 1) = "Delta"

  word_list(TOTAL_DELTA_N, 0) = "Total Delta"
  word_list(TOTAL_DELTA_N, 1) = "Total Delta"

  word_list(MINIMUM_N, 0) = "Minimum"
  word_list(MINIMUM_N, 1) = "Mínimo"

  word_list(EST_BALANCE_N, 0) = "Est Balance"
  word_list(EST_BALANCE_N, 1) = "Balance estimado"

  word_list(HIGH_LOW_N, 0) = "High/Low"
  word_list(HIGH_LOW_N, 1) = "Alto/Bajo"

  word_list(BEGIN_END_N, 0) = "Begin/End"
  word_list(BEGIN_END_N, 1) = "Inicio/Fin"

  word_list(DETAILED_SUMMARY_N, 0) = "Detailed Summary"
  word_list(DETAILED_SUMMARY_N, 1) = "Resumen Detallado"

  word_list(REGISTER_N, 0) = "Register"
  word_list(REGISTER_N, 1) = "Registro"

  word_list(PURCHASE_INFORMATION_N, 0) = "Purchase Information"
  word_list(PURCHASE_INFORMATION_N, 1) = "Datos de la Adquisición"

  word_list(RUN_CHECK2CHECK_N, 0) = "Run Check2Check"
  word_list(RUN_CHECK2CHECK_N, 1) = "Iniciar Check2Check"

  word_list(CODE_N, 0) = "Code"
  word_list(CODE_N, 1) = "Código"

  word_list(NOT_ENOUGH_INFORMATION_N, 0) = "Not enough information entered"
  word_list(NOT_ENOUGH_INFORMATION_N, 1) = "Información no suficiente"

  word_list(TOTAL_N, 0) = "Total"
  word_list(TOTAL_N, 1) = "Total"

  word_list(QTY_N, 0) = "Qty"
  word_list(QTY_N, 1) = "Qty"

  word_list(TO_N, 0) = "to"
  word_list(TO_N, 1) = "to"

  word_list(QTY_INC_N, 0) = "Qty Inc"
  word_list(QTY_INC_N, 1) = "Qty Inc"

  word_list(QTY_EXP_N, 0) = "Qty Exp"
  word_list(QTY_EXP_N, 1) = "Qty Exp"

  word_list(INCOME_N, 0) = "Inc"
  word_list(INCOME_N, 1) = "Inc"

  word_list(EXPENSE_N, 0) = "Exp"
  word_list(EXPENSE_N, 1) = "Exp"

  word_list(CHARACTERS_LEFT_N, 0) = "Characters Left"
  word_list(CHARACTERS_LEFT_N, 1) = "Caracteres dejados"

  word_list(NEXT_MONTH_TOP_N, 0) = "Next Month Top"
  word_list(NEXT_MONTH_TOP_N, 1) = "Tapa siguiente del mes"

  word_list(PREVIOUS_MONTH_BOTTOM_N, 0) = "Previous Month Bottom"
  word_list(PREVIOUS_MONTH_BOTTOM_N, 1) = "fondo anterior del mes"

  
  
  
  
  
  word_list(CHANGE_N, 0) = "Change"
  word_list(CHANGE_N, 1) = "?????"

  word_list(ALL_CARDS_N, 0) = "All Cards"
  word_list(ALL_CARDS_N, 1) = "?????"

  word_list(EST_BAL_N, 0) = "Est Bal"
  word_list(EST_BAL_N, 1) = "?????"

  word_list(CLEARED_N, 0) = "Cleared"
  word_list(CLEARED_N, 1) = "Aclaradas"

  word_list(FINISHED_N, 0) = "Finished"
  word_list(FINISHED_N, 1) = "Acabado"

  word_list(SORT_BY_N, 0) = "Sort by"
  word_list(SORT_BY_N, 1) = "?????"

  word_list(PROMPT_FOR_MOVE_N, 0) = "Prompt for Move"
  word_list(PROMPT_FOR_MOVE_N, 1) = "?????"

  word_list(PROMPT_FOR_DELETE_N, 0) = "Prompt for Delete"
  word_list(PROMPT_FOR_DELETE_N, 1) = "?????"

  word_list(SHOW_NAMES_IN_COLORS_N, 0) = "Show names in colors"
  word_list(SHOW_NAMES_IN_COLORS_N, 1) = "?????"

  word_list(AUTO_CHECK_DONE_ON_AMOUNT_N, 0) = "Auto check 'Done' on amount entry"
  word_list(AUTO_CHECK_DONE_ON_AMOUNT_N, 1) = "?????"

  word_list(AUTO_INSERT_NEW_LINE_N, 0) = "Auto insert new line on transaction entry"
  word_list(AUTO_INSERT_NEW_LINE_N, 1) = "?????"

  word_list(SHOW_OVERRIDE_COLUMNS_N, 0) = "Show override columns at startup"
  word_list(SHOW_OVERRIDE_COLUMNS_N, 1) = "?????"

  word_list(SHOW_SPLASH_SCREEN_N, 0) = "Show splash screen at startup"
  word_list(SHOW_SPLASH_SCREEN_N, 1) = "?????"

  word_list(PROMPT_FOR_PASTE_NOTES_N, 0) = "Prompt for paste notes"
  word_list(PROMPT_FOR_PASTE_NOTES_N, 1) = "?????"

  word_list(AUTO_NEGATIVE_NUMBERS_N, 0) = "Auto negative numbers in amount column"
  word_list(AUTO_NEGATIVE_NUMBERS_N, 1) = "?????"

  word_list(AUTO_CHECK_DONE_ON_CHECK_N, 0) = "Auto check 'Done' on check entry"
  word_list(AUTO_CHECK_DONE_ON_CHECK_N, 1) = "?????"

  word_list(SAVE_RECOVERY_FILE_N, 0) = "Save recovery file"
  word_list(SAVE_RECOVERY_FILE_N, 1) = "?????"

  word_list(SAME_N, 0) = "Same"
  word_list(SAME_N, 1) = "?????"

  word_list(CHANGE_PASSWORD_N, 0) = "Change Password"
  word_list(CHANGE_PASSWORD_N, 1) = "?????"

  word_list(FONT_SIZE_NOTES_N, 0) = "Font size for notes"
  word_list(FONT_SIZE_NOTES_N, 1) = "?????"

  word_list(GENERAL_N, 0) = "General"
  word_list(GENERAL_N, 1) = "?????"

  word_list(STATUS_COLUMN_N, 0) = "Status Column"
  word_list(STATUS_COLUMN_N, 1) = "?????"

  word_list(CHECK_NUMBER_COLUMN_N, 0) = "Check Number Column"
  word_list(CHECK_NUMBER_COLUMN_N, 1) = "?????"

  word_list(CLEARED_COLUMN_N, 0) = "Cleared Column"
  word_list(CLEARED_COLUMN_N, 1) = "?????"

  word_list(PASSWORD_N, 0) = "Password"
  word_list(PASSWORD_N, 1) = "?????"

  word_list(ENTER_PASSWORD_N, 0) = "Enter Password"
  word_list(ENTER_PASSWORD_N, 1) = "?????"

  word_list(ENTER_NEW_PASSWORD_N, 0) = "Enter New Password Twice"
  word_list(ENTER_NEW_PASSWORD_N, 1) = "?????"

  word_list(AUTO_LOAD_LAST_FILE_N, 0) = "Auto load last file"
  word_list(AUTO_LOAD_LAST_FILE_N, 1) = "?????"

  
  
  
  
  
  word_list(NO_TRANSACTIONS_COPIED_N, 0) = "No transactions copied"
  word_list(NO_TRANSACTIONS_COPIED_N, 1) = "?????"

  word_list(NO_TRANSACTIONS_CUT_N, 0) = "No transactions cut"
  word_list(NO_TRANSACTIONS_CUT_N, 1) = "?????"

  word_list(NO_TRANSACTIONS_TO_CUT_N, 0) = "No transactions to cut"
  word_list(NO_TRANSACTIONS_TO_CUT_N, 1) = "?????"

  word_list(CUT_ALL_SELECTED_Q_N, 0) = "Cut all selected?"
  word_list(CUT_ALL_SELECTED_Q_N, 1) = "?????"

  word_list(TRANSACTIONS_FOR_N, 0) = "transactions for"
  word_list(TRANSACTIONS_FOR_N, 1) = "?????"

  word_list(NO_TRANSACTIONS_TO_COPY_N, 0) = "No transactions to copy"
  word_list(NO_TRANSACTIONS_TO_COPY_N, 1) = "?????"

  word_list(CUT_ALL_TAGGED_Q_N, 0) = "Cut all tagged?"
  word_list(CUT_ALL_TAGGED_Q_N, 1) = "?????"

  word_list(DELETE_TRANSACTION_Q_N, 0) = "Delete Transaction?"
  word_list(DELETE_TRANSACTION_Q_N, 1) = "?????"

  word_list(CUT_ALL_TRANSACTIONS_FOR_N, 0) = "Cut all transactions for"
  word_list(CUT_ALL_TRANSACTIONS_FOR_N, 1) = "?????"
  
  word_list(MOVE_TRANSACTION_Q_N, 0) = "Move Transaction?"
  word_list(MOVE_TRANSACTION_Q_N, 1) = "?????"

  word_list(CANT_ASSIGN_A_CHECK_NUMBER_TO_A_DEPOSIT_N, 0) = "Can't assign a check number to a deposit"
  word_list(CANT_ASSIGN_A_CHECK_NUMBER_TO_A_DEPOSIT_N, 1) = "?????"

  word_list(CHECK_NUMBER_ERROR_N, 0) = "Check Number Error"
  word_list(CHECK_NUMBER_ERROR_N, 1) = "?????"

  word_list(SAVE_DATABASE_Q_N, 0) = "Save Database?"
  word_list(SAVE_DATABASE_Q_N, 1) = "?????"

  word_list(ERROR_N, 0) = "Error"
  word_list(ERROR_N, 1) = "?????"

  word_list(PASTE_ALL_TRANSACTIONS_TO_N, 0) = "Paste all transactions to"
  word_list(PASTE_ALL_TRANSACTIONS_TO_N, 1) = "?????"

  word_list(PASTE_NOTES_Q_N, 0) = "Paste notes?"
  word_list(PASTE_NOTES_Q_N, 1) = "?????"

  word_list(ADJUSTED_DATES_TO_MATCH_THE_CURRENT_MONTH_N, 0) = "Adjusted dates to match the current month"
  word_list(ADJUSTED_DATES_TO_MATCH_THE_CURRENT_MONTH_N, 1) = "?????"

  word_list(CANT_PASTE_TO_THE_SAME_MONTH_AND_YEAR_N, 0) = "Can't paste to the same month and year"
  word_list(CANT_PASTE_TO_THE_SAME_MONTH_AND_YEAR_N, 1) = "?????"

  word_list(PASTE_NOTES_Q_N, 0) = "Paste notes?"
  word_list(PASTE_NOTES_Q_N, 1) = "?????"

  word_list(FILE_NOT_FOUND_N, 0) = "File not found."
  word_list(FILE_NOT_FOUND_N, 1) = "?????"

  word_list(FILE_EXISTS_OVERWRITE_Q_N, 0) = "File Exists. Overwrite?"
  word_list(FILE_EXISTS_OVERWRITE_Q_N, 1) = "?????"

  word_list(OVERWRITE_Q_N, 0) = "Overwrite?"
  word_list(OVERWRITE_Q_N, 1) = "?????"

  word_list(ERROR_IN_DELETING_WRITING_FILE_N, 0) = "Error in Deleting/Writing file"
  word_list(ERROR_IN_DELETING_WRITING_FILE_N, 1) = "?????"

  word_list(INVALID_NUMBER_ENTERED_N, 0) = "Invalid number entered"
  word_list(INVALID_NUMBER_ENTERED_N, 1) = "?????"

  word_list(INVALID_TAG_NUMBER_N, 0) = "Invalid tag number"
  word_list(INVALID_TAG_NUMBER_N, 1) = "?????"

  word_list(CHANGING_THIS_MAY_CAUSE_A_PROBLEM_WHEN_YOU_RECONCILE_Q_N, 0) = "Changing this may cause a problem when you reconcile next time. Are you sure you want to change it?"
  word_list(CHANGING_THIS_MAY_CAUSE_A_PROBLEM_WHEN_YOU_RECONCILE_Q_N, 1) = "?????"

  word_list(INVALID_KEY_N, 0) = "Invalid key - Only space, * and X are permitted"
  word_list(INVALID_KEY_N, 1) = "?????"

  word_list(INVALID_NUMBER_ENTERED_N, 0) = "Invalid number entered"
  word_list(INVALID_NUMBER_ENTERED_N, 1) = "?????"

  word_list(CALCULATOR_NOT_FOUND_N, 0) = "Calculator not found."
  word_list(CALCULATOR_NOT_FOUND_N, 1) = "?????"

  word_list(SAVE_CHANGES_Q_N, 0) = "Save Changes?"
  word_list(SAVE_CHANGES_Q_N, 1) = "?????"

  word_list(DELETE_ALL_CREDIT_CARDS_AND_INFORMATION_Q_N, 0) = "Delete ALL Credit Cards and Information?"
  word_list(DELETE_ALL_CREDIT_CARDS_AND_INFORMATION_Q_N, 1) = "?????"

  word_list(ARE_YOU_ABSOLUTELY_SURE_Q_N, 0) = "Are You Absolutely Sure?"
  word_list(ARE_YOU_ABSOLUTELY_SURE_Q_N, 1) = "?????"

  word_list(ALL_CREDIT_CARDS_HAVE_BEEN_DELETED_N, 0) = "All credit cards have been deleted"
  word_list(ALL_CREDIT_CARDS_HAVE_BEEN_DELETED_N, 1) = "?????"


  word_list(DELETE_Q_N, 0) = "Delete?"
  word_list(DELETE_Q_N, 1) = "?????"

  word_list(NEW_TRANSACTION_ENTERED_N, 0) = "New transaction entered"
  word_list(NEW_TRANSACTION_ENTERED_N, 1) = "?????"

  word_list(MUST_HAVE_A_TRANSACTION_NAME_N, 0) = "Must have a transaction name"
  word_list(MUST_HAVE_A_TRANSACTION_NAME_N, 1) = "?????"

  word_list(MUST_SELECT_A_CREDIT_CARD_N, 0) = "Must select a credit card"
  word_list(MUST_SELECT_A_CREDIT_CARD_N, 1) = "?????"

  word_list(SELECT_CARD_N, 0) = "Select card"
  word_list(SELECT_CARD_N, 1) = "?????"

  word_list(ERRORS_IN_CT_DATABASE_N, 0) = "errors in CT database, 1st="
  word_list(ERRORS_IN_CT_DATABASE_N, 1) = "?????"

  word_list(ADD_NEW_CARD_Q_N, 0) = "Add new card?"
  word_list(ADD_NEW_CARD_Q_N, 1) = "?????"

  word_list(SORRY_NO_SPACE_FOR_CARDS_N, 0) = "Sorry, no space for cards"
  word_list(SORRY_NO_SPACE_FOR_CARDS_N, 1) = "?????"

  word_list(ARE_YOU_SURE_N, 0) = "Are you sure?"
  word_list(ARE_YOU_SURE_N, 1) = "?????"

  word_list(INVALID_ENTRY_N, 0) = "Invalid entry"
  word_list(INVALID_ENTRY_N, 1) = "?????"

  word_list(PASSWORDS_DONT_MATCH_N, 0) = "Passwords don't match"
  word_list(PASSWORDS_DONT_MATCH_N, 1) = "?????"

  word_list(DELETE_QUICK_SAVE_ACCOUNT_Q_N, 0) = "Delete Quick Save account?"
  word_list(DELETE_QUICK_SAVE_ACCOUNT_Q_N, 1) = "?????"

  word_list(DELETE_QUICK_SAVE_ACCOUNT_N, 0) = "Delete Quick Save account"
  word_list(DELETE_QUICK_SAVE_ACCOUNT_N, 1) = "?????"

  word_list(ARE_YOU_SURE_Q_QUICK_SAVE_ACCOUNT_IS_NOT_ZERO_N, 0) = "Are you sure? Quick Save account balance is not zero."
  word_list(ARE_YOU_SURE_Q_QUICK_SAVE_ACCOUNT_IS_NOT_ZERO_N, 1) = "?????"

  word_list(INVALID_NUMBER_N, 0) = "Invalid number"
  word_list(INVALID_NUMBER_N, 1) = "?????"

  word_list(SORRY_NOT_ENOUGH_MONEY_IN_THAT_ACCOUNT_N, 0) = "Sorry, not enough money in that account. Maximum allowed is"
  word_list(SORRY_NOT_ENOUGH_MONEY_IN_THAT_ACCOUNT_N, 1) = "?????"

  word_list(SAVE_QUICK_ACCOUNT_DATA_Q_N, 0) = "Save Quick Account Data?"
  word_list(SAVE_QUICK_ACCOUNT_DATA_Q_N, 1) = "?????"

  word_list(SAVE_QUICK_ACCOUNT_DATA_N, 0) = "Save Quick Account Data"
  word_list(SAVE_QUICK_ACCOUNT_DATA_N, 1) = "?????"

  word_list(INVALID_AMOUNT_N, 0) = "Invalid amount"
  word_list(INVALID_AMOUNT_N, 1) = "?????"

  word_list(DO_YOU_WANT_TO_FINALIZE_Q_N, 0) = "Do you want to finalize the reconciliation process now?"
  word_list(DO_YOU_WANT_TO_FINALIZE_Q_N, 1) = "?????"

  word_list(THE_DIFFERENCE_IS_NOT_ZERO_Q_N, 0) = "The difference is not $0.00. Are you sure you want to finish the reconciliation process now?"
  word_list(THE_DIFFERENCE_IS_NOT_ZERO_Q_N, 1) = "?????"

  word_list(SUCCESS_THANK_YOU_FOR_REGISTERING_N, 0) = "Success. Thank you for registering "
  word_list(SUCCESS_THANK_YOU_FOR_REGISTERING_N, 1) = "?????"

  word_list(SORRY_NOT_A_VALID_NAME_CODE_N, 0) = "Sorry, not a valid Name / Code."
  word_list(SORRY_NOT_A_VALID_NAME_CODE_N, 1) = "?????"

  word_list(TRANSACTIONS_FOR_N, 0) = "transactions for"
  word_list(TRANSACTIONS_FOR_N, 1) = "?????"

  word_list(SAVE_Q_N, 0) = "Save?"
  word_list(SAVE_Q_N, 1) = "?????"

  word_list(ADD_Q_N, 0) = "Add?"
  word_list(ADD_Q_N, 1) = "?????"

  word_list(DELETE_Q_N, 0) = "Delete?"
  word_list(DELETE_Q_N, 1) = "?????"

  word_list(SELECT_THIS_CARD_Q_N, 0) = "Select this card?"
  word_list(SELECT_THIS_CARD_Q_N, 1) = "?????"

  word_list(CHANGE_FINISHED_STATUS_Q_N, 0) = "Change finished status?"
  word_list(CHANGE_FINISHED_STATUS_Q_N, 1) = "?????"

  word_list(DELETE_ACCOUNT_Q_N, 0) = "Delete account?"
  word_list(DELETE_ACCOUNT_Q_N, 1) = "?????"

  word_list(WARNING_QUICK_SAVE_ACCOUNT_BALANCE_IS_NOT_ZERO_N, 0) = "Warning, Quick Save account balance not zero."
  word_list(WARNING_QUICK_SAVE_ACCOUNT_BALANCE_IS_NOT_ZERO_N, 1) = "?????"

  word_list(CANCEL_RECONCILIATION_Q_N, 0) = "Cancel reconcilation?"
  word_list(CANCEL_RECONCILIATION_Q_N, 1) = "?????"

  word_list(SUCCESS_RECONCILIATION_COMPLETE_N, 0) = "Success - Reconciliation Complete"
  word_list(SUCCESS_RECONCILIATION_COMPLETE_N, 1) = "?????"

  word_list(CAUTION_RECONCILIATION_NOT_COMPLETE_N, 0) = "Caution - Reconcilation Not Complete"
  word_list(CAUTION_RECONCILIATION_NOT_COMPLETE_N, 1) = "?????"

  
  
  
  
  word_list(CREDIT_CARD_TRANSACTIONS_N, 0) = "Credit Card Transactions"
  word_list(CREDIT_CARD_TRANSACTIONS_N, 1) = "?????"

  word_list(CARD_N, 0) = "Card"
  word_list(CARD_N, 1) = "?????"

  word_list(DELETE_CARD_N, 0) = "Delete Card"
  word_list(DELETE_CARD_N, 1) = "?????"

  word_list(DELETE_ALL_CARDS_N, 0) = "Delete All Cards"
  word_list(DELETE_ALL_CARDS_N, 1) = "?????"

  word_list(SELECT_THIS_CARD_N, 0) = "Select This Card"
  word_list(SELECT_THIS_CARD_N, 1) = "?????"

  word_list(DELETE_THIS_CARD_N, 0) = "Delete This Card"
  word_list(DELETE_THIS_CARD_N, 1) = "?????"

  word_list(SELECT_N, 0) = "Select"
  word_list(SELECT_N, 1) = "?????"

  word_list(NAME_ADDRESS_PHONE_NUMBERS_N, 0) = "Name / Address / Phone Numbers / Credit Card Limits"
  word_list(NAME_ADDRESS_PHONE_NUMBERS_N, 1) = "?????"

  word_list(DATE_DUE_N, 0) = "Date Due"
  word_list(DATE_DUE_N, 1) = "?????"

  word_list(CREDIT_CARD_INFORMATION_N, 0) = "Credit Card Information"
  word_list(CREDIT_CARD_INFORMATION_N, 1) = "?????"

  word_list(ACCOUNT_N, 0) = "Account"
  word_list(ACCOUNT_N, 1) = "?????"

  word_list(FIRST_TRANSACTION_N, 0) = "First Transaction"
  word_list(FIRST_TRANSACTION_N, 1) = "?????"

  word_list(LAST_TRANSACTION_N, 0) = "Last Transaction"
  word_list(LAST_TRANSACTION_N, 1) = "?????"

  word_list(TOTAL_TRANSACTIONS_N, 0) = "Total Transactions"
  word_list(TOTAL_TRANSACTIONS_N, 1) = "?????"

  word_list(NEXT_CHECK_NUMBER_N, 0) = "Next Check Number"
  word_list(NEXT_CHECK_NUMBER_N, 1) = "?????"

  word_list(VERSION_N, 0) = "Version"
  word_list(VERSION_N, 1) = "?????"

  word_list(TOTAL_NOTES_N, 0) = "Total Notes"
  word_list(TOTAL_NOTES_N, 1) = "?????"

  word_list(TOTAL_CARDTRAKS_N, 0) = "Total Cardtraks"
  word_list(TOTAL_CARDTRAKS_N, 1) = "Total Rastreo-Tarjetas"

  word_list(PLAY_SOUNDS_N, 0) = "Play sounds"
  word_list(PLAY_SOUNDS_N, 1) = "?????"

  word_list(PASTE_N, 0) = "Paste"
  word_list(PASTE_N, 1) = "?????"

  word_list(PASTE_OPTIONS_N, 0) = "Paste Options"
  word_list(PASTE_OPTIONS_N, 1) = "?????"

  word_list(PASTE_TAGS_OPTIONS_N, 0) = "Paste Tags Options"
  word_list(PASTE_TAGS_OPTIONS_N, 1) = "?????"

End Sub


