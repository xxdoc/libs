Option Explicit

Public Enum activationfunc_enum
  FANN_LINEAR = 0
  FANN_THRESHOLD = 1
  FANN_THRESHOLD_SYMMETRIC = 2
  FANN_SIGMOID = 3
  FANN_SIGMOID_STEPWISE = 4
  FANN_SIGMOID_SYMMETRIC = 5
  FANN_SIGMOID_SYMMETRIC_STEPWISE = 6
  FANN_GAUSSIAN = 7
  FANN_GAUSSIAN_SYMMETRIC = 8
  FANN_GAUSSIAN_STEPWISE = 9
  FANN_ELLIOT = 10
  FANN_ELLIOT_SYMMETRIC = 11
  FANN_LINEAR_PIECE = 12
  FANN_LINEAR_PIECE_SYMMETRIC = 13
  FANN_SIN_SYMMETRIC = 14
  FANN_COS_SYMMETRIC = 15
  FANN_SIN = 16
  FANN_COS = 17
End Enum

Public Enum stopfunc_enum
  FANN_STOP_MSE
  FANN_STOPFUNC_BIT
End Enum

Public Declare Function fann_create_standard_array Lib "fanndouble" Alias "_fann_create_standard_array@8" (ByVal num_layers As Long, layers As Long) As Long
 
Public Declare Function fann_read_train_from_file Lib "fanndouble" Alias "_fann_read_train_from_file@4" (ByVal Filename As String) As Long
Public Declare Sub fann_set_activation_steepness_hidden Lib "fanndouble" Alias "_fann_set_activation_steepness_hidden@12" (ByVal ann As Long, ByVal steepness As Double)
Public Declare Sub fann_set_activation_steepness_output Lib "fanndouble" Alias "_fann_set_activation_steepness_output@12" (ByVal ann As Long, ByVal steepness As Double)
Public Declare Sub fann_set_activation_function_hidden Lib "fanndouble" Alias "_fann_set_activation_function_hidden@8" (ByVal ann As Long, ByVal activationfunc As activationfunc_enum)
Public Declare Sub fann_set_activation_function_output Lib "fanndouble" Alias "_fann_set_activation_function_output@8" (ByVal ann As Long, ByVal activationfunc As activationfunc_enum)

Public Declare Sub fann_set_train_stop_function Lib "fanndouble" Alias "_fann_set_train_stop_function@8" (ByVal ann As Long, ByVal stopfunc As stopfunc_enum)
Public Declare Sub fann_set_bit_fail_limit Lib "fanndouble" Alias "_fann_set_bit_fail_limit@12" (ByVal ann As Long, ByVal bit_fail_limit As Double)

Public Declare Sub fann_init_weights Lib "fanndouble" Alias "_fann_init_weights@8" (ByVal ann As Long, ByVal train_data As Long)
Public Declare Sub fann_train_on_data Lib "fanndouble" Alias "_fann_train_on_data@20" (ByVal ann As Long, ByVal train_data As Long, _
                   ByVal max_epochs As Long, ByVal epochs_between_reports As Long, ByVal desired_error As Single)

Public Declare Function fann_run Lib "fanndouble" Alias "_fann_run@8" (ByVal ann As Long, inputs As Double) As Long
Public Declare Sub fann_destroy Lib "fanndouble" Alias "_fann_destroy@4" (ByVal ann As Long)

'**** x86 __stdcall function-names with name decorations (as copied from depends.exe-GUI)
'_fann_cascadetrain_on_data@20
'_fann_cascadetrain_on_file@20
'_fann_clear_scaling_params@4
'_fann_copy@4
'_fann_create_from_file@4
'_fann_create_shortcut_array@8
'_fann_create_sparse_array@12
'_fann_create_standard_array@8
'_fann_create_train@12
'_fann_create_train_array@20
'_fann_create_train_from_callback@16
'_fann_create_train_pointer_array@20
'_fann_descale_input@8
'_fann_descale_output@8
'_fann_descale_train@8
'_fann_destroy@4
'_fann_destroy_train@4
'_fann_duplicate_train_data@4
'_fann_get_activation_function@12
'_fann_get_activation_steepness@12
'_fann_get_bias_array@8
'_fann_get_bit_fail@4
'_fann_get_bit_fail_limit@4
'_fann_get_callback@4
'_fann_get_cascade_activation_functions@4
'_fann_get_cascade_activation_functions_count@4
'_fann_get_cascade_activation_steepnesses@4
'_fann_get_cascade_activation_steepnesses_count@4
'_fann_get_cascade_candidate_change_fraction@4
'_fann_get_cascade_candidate_limit@4
'_fann_get_cascade_candidate_stagnation_epochs@4
'_fann_get_cascade_max_cand_epochs@4
'_fann_get_cascade_max_out_epochs@4
'_fann_get_cascade_min_cand_epochs@4
'_fann_get_cascade_min_out_epochs@4
'_fann_get_cascade_num_candidate_groups@4
'_fann_get_cascade_num_candidates@4
'_fann_get_cascade_output_change_fraction@4
'_fann_get_cascade_output_stagnation_epochs@4
'_fann_get_cascade_weight_multiplier@4
'_fann_get_connection_array@8
'_fann_get_connection_rate@4
'_fann_get_errno@4
'_fann_get_errstr@4
'_fann_get_layer@8
'_fann_get_layer_array@8
'_fann_get_learning_momentum@4
'_fann_get_learning_rate@4
'_fann_get_MSE@4
'_fann_get_network_type@4
'_fann_get_neuron@12
'_fann_get_neuron_layer@12
'_fann_get_num_input@4
'_fann_get_num_layers@4
'_fann_get_num_output@4
'_fann_get_quickprop_decay@4
'_fann_get_quickprop_mu@4
'_fann_get_rprop_decrease_factor@4
'_fann_get_rprop_delta_max@4
'_fann_get_rprop_delta_min@4
'_fann_get_rprop_delta_zero@4
'_fann_get_rprop_increase_factor@4
'_fann_get_sarprop_step_error_shift@4
'_fann_get_sarprop_step_error_threshold_factor@4
'_fann_get_sarprop_temperature@4
'_fann_get_sarprop_weight_decay_shift@4
'_fann_get_total_connections@4
'_fann_get_total_neurons@4
'_fann_get_train_error_function@4
'_fann_get_train_input@8
'_fann_get_train_output@8
'_fann_get_train_stop_function@4
'_fann_get_training_algorithm@4
'_fann_get_user_data@4
'_fann_init_weights@8
'_fann_length_train_data@4
'_fann_merge_train_data@8
'_fann_num_input_train_data@4
'_fann_num_output_train_data@4
'_fann_print_connections@4
'_fann_print_error@4
'_fann_print_parameters@4
'_fann_randomize_weights@20
'_fann_read_train_from_file@4
'_fann_reset_errno@4
'_fann_reset_errstr@4
'_fann_reset_MSE@4
'_fann_run@8
'_fann_save@8
'_fann_save_to_fixed@8
'_fann_save_train@8
'_fann_save_train_to_fixed@12
'_fann_scale_data_to_range@44
'_fann_scale_input@8
'_fann_scale_input_train_data@20
'_fann_scale_output@8
'_fann_scale_output_train_data@20
'_fann_scale_train@8
'_fann_scale_train_data@20
'_fann_set_activation_function@16
'_fann_set_activation_function_hidden@8
'_fann_set_activation_function_layer@12
'_fann_set_activation_function_output@8
'_fann_set_activation_steepness@20
'_fann_set_activation_steepness_hidden@12
'_fann_set_activation_steepness_layer@16
'_fann_set_activation_steepness_output@12
'_fann_set_bit_fail_limit@12
'_fann_set_callback@8
'_fann_set_cascade_activation_functions@12
'_fann_set_cascade_activation_steepnesses@12
'_fann_set_cascade_candidate_change_fraction@8
'_fann_set_cascade_candidate_limit@12
'_fann_set_cascade_candidate_stagnation_epochs@8
'_fann_set_cascade_max_cand_epochs@8
'_fann_set_cascade_max_out_epochs@8
'_fann_set_cascade_min_cand_epochs@8
'_fann_set_cascade_min_out_epochs@8
'_fann_set_cascade_num_candidate_groups@8
'_fann_set_cascade_output_change_fraction@8
'_fann_set_cascade_output_stagnation_epochs@8
'_fann_set_cascade_weight_multiplier@12
'_fann_set_error_log@8
'_fann_set_input_scaling_params@16
'_fann_set_learning_momentum@8
'_fann_set_learning_rate@8
'_fann_set_output_scaling_params@16
'_fann_set_quickprop_decay@8
'_fann_set_quickprop_mu@8
'_fann_set_rprop_decrease_factor@8
'_fann_set_rprop_delta_max@8
'_fann_set_rprop_delta_min@8
'_fann_set_rprop_delta_zero@8
'_fann_set_rprop_increase_factor@8
'_fann_set_sarprop_step_error_shift@8
'_fann_set_sarprop_step_error_threshold_factor@8
'_fann_set_sarprop_temperature@8
'_fann_set_sarprop_weight_decay_shift@8
'_fann_set_scaling_params@24
'_fann_set_train_error_function@8
'_fann_set_train_stop_function@8
'_fann_set_training_algorithm@8
'_fann_set_user_data@8
'_fann_set_weight@20
'_fann_set_weight_array@12
'_fann_shuffle_train_data@4
'_fann_subset_train_data@12
'_fann_test@12
'_fann_test_data@8
'_fann_train@12
'_fann_train_epoch@8
'_fann_train_epoch_batch_parallel@12
'_fann_train_epoch_incremental_mod@8
'_fann_train_epoch_irpropm_parallel@12
'_fann_train_epoch_quickprop_parallel@12
'_fann_train_epoch_sarprop_parallel@12
'_fann_train_on_data@20
'_fann_train_on_file@20
 

'***** working declares (used in my tests with a __stdcall-compile with the tinycc-compiler, without name-decorations)

'Public Declare Function fann_create_standard_array Lib "doublefann" (ByVal num_layers As Long, layers As Long) As Long
'
'Public Declare Function fann_read_train_from_file Lib "doublefann" (ByVal Filename As String) As Long
'Public Declare Sub fann_set_activation_steepness_hidden Lib "doublefann" (ByVal ann As Long, ByVal steepness As Double)
'Public Declare Sub fann_set_activation_steepness_output Lib "doublefann" (ByVal ann As Long, ByVal steepness As Double)
'Public Declare Sub fann_set_activation_function_hidden Lib "doublefann" (ByVal ann As Long, ByVal activationfunc As activationfunc_enum)
'Public Declare Sub fann_set_activation_function_output Lib "doublefann" (ByVal ann As Long, ByVal activationfunc As activationfunc_enum)
'
'Public Declare Sub fann_set_train_stop_function Lib "doublefann" (ByVal ann As Long, ByVal stopfunc As stopfunc_enum)
'Public Declare Sub fann_set_bit_fail_limit Lib "doublefann" (ByVal ann As Long, ByVal bit_fail_limit As Double)
'
'Public Declare Sub fann_init_weights Lib "doublefann" (ByVal ann As Long, ByVal train_data As Long)
'Public Declare Sub fann_train_on_data Lib "doublefann" (ByVal ann As Long, ByVal train_data As Long, _
'                   ByVal max_epochs As Long, ByVal epochs_between_reports As Long, ByVal desired_error As Single)
'
'Public Declare Function fann_run Lib "doublefann" (ByVal ann As Long, inputs As Double) As Long
'Public Declare Sub fann_destroy Lib "doublefann" (ByVal ann As Long)


