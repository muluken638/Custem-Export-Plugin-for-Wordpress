<?php
// pages/excel_export_form.php

if (!defined('ABSPATH')) exit; // security check
?>
<div class="wrap" style=" margin:20px ">
    <h1 style="color:#21759b; font-weight:bold; text-align:center; margin-bottom:30px;">Custom Excel Export</h1>
    
    <form method="get" action="<?php echo admin_url('admin-post.php'); ?>" style="background:#ffffff; padding:30px; border-radius:12px; box-shadow: 0 6px 18px rgba(0,0,0,0.1); border:1px solid #ddd; transition: all 0.3s ease;">
        <input type="hidden" name="action" value="cee_export_excel">

        <h2 style="color:#333; margin-bottom:20px; border-bottom:2px solid #21759b; padding-bottom:5px;">Filter Applicants</h2>

        <div style="display:flex; flex-wrap:wrap; gap:20px; margin-bottom:20px;">
            <div style="flex:1; min-width:200px;">
                <label style="display:block; font-weight:bold; color:#555; margin-bottom:5px;">Date From:</label>
                <input type="date" name="date_from" style="width:100%; padding:8px 10px; border-radius:6px; border:1px solid #ccc; transition: all 0.2s;"/>
            </div>
            <div style="flex:1; min-width:200px;">
                <label style="display:block; font-weight:bold; color:#555; margin-bottom:5px;">Date To:</label>
                <input type="date" name="date_to" style="width:100%; padding:8px 10px; border-radius:6px; border:1px solid #ccc; transition: all 0.2s;"/>
            </div>
        </div>

        <div style="display:flex; flex-wrap:wrap; gap:20px;">
            <?php
            $fields = [
                'Gender' => ['name'=>'gender','options'=>$genders],
                'Educational Level' => ['name'=>'educational_level','options'=>$educational_level],
                'Marital Status' => ['name'=>'marital_status','options'=>$marital_statuses],
                'Job Category' => ['name'=>'job_category','options'=>$job_categories],
                'Job Title' => ['name'=>'job_title','options'=>$job_titles],
                'Job Type' => ['name'=>'job_type','options'=>$job_types],
                'Job Location' => ['name'=>'job_location','options'=>$job_locations],
            ];

            foreach($fields as $label=>$data){
                echo '<div style="flex:1; min-width:220px;">
                        <label style="display:block; font-weight:bold; color:#555; margin-bottom:5px;">'.$label.':</label>
                        <select name="'.$data['name'].'" style="width:100%; padding:8px 10px; border-radius:6px; border:1px solid #ccc; transition: all 0.2s;">
                            <option value="">--All--</option>';
                foreach($data['options'] as $opt){
                    echo '<option value="'.esc_attr($opt).'">'.esc_html($opt).'</option>';
                }
                echo '</select>
                      </div>';
            }
            ?>
        </div>

        <div style="margin-top:30px; text-align:end;">
            <input type="submit" class="button button-primary" value="Download Excel" 
                style="padding:12px 30px; background:#21759b; border:none; border-radius:8px; font-weight:bold; cursor:pointer; transition: all 0.3s;">
        </div>
    </form>

    <style>
        form input[type="date"]:focus,
        form select:focus {
            border-color: #21759b;
            box-shadow: 0 0 5px rgba(33,117,155,0.5);
            outline: none;
        }

        form input[type="submit"]:hover {
            background:#1b4f73;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        }

        form:hover {
            box-shadow: 0 8px 22px rgba(0,0,0,0.15);
        }
    </style>
</div>
