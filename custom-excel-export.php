<?php
/*
Plugin Name: Custom Excel Export
Description: Admin-only plugin to export Simple Job Board applicants to Excel with dynamic fields, filters, and saved copies. Dropdowns are dynamically populated from the database.
Version: 2.5
Author: Muluken zeleke
*/

if (!defined('ABSPATH')) exit;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;

// Load PhpSpreadsheet
require_once plugin_dir_path(__FILE__) . 'vendor/autoload.php';

// Admin menu
add_action('admin_menu', function() {
    add_menu_page(
        'Custom Excel Export',
        'Excel Export',
        'manage_options',
        'custom-excel-export',
        'cee_admin_page',
        'dashicons-download',
        100
    );
});

// Admin page with dynamic filters
function cee_admin_page() {
    global $wpdb;

    // Fetch dynamic genders
    $genders = [];
    $gender_results = $wpdb->get_results("SELECT DISTINCT meta_value FROM {$wpdb->postmeta} WHERE meta_key='jobapp_gender'");
    foreach ($gender_results as $r) {
        $value = maybe_unserialize($r->meta_value);
        if (is_array($value)) {
            foreach ($value as $v) $genders[] = $v;
        } else {
            $genders[] = $value;
        }
    }
    $genders = array_unique($genders);
    sort($genders);

    // Fetch dynamic marital statuses
    $marital_statuses = [];
    $status_results = $wpdb->get_results("SELECT DISTINCT meta_value FROM {$wpdb->postmeta} WHERE meta_key='jobapp_marital_status'");
    foreach ($status_results as $r) {
        $value = maybe_unserialize($r->meta_value);
        if (is_array($value)) {
            foreach ($value as $v) $marital_statuses[] = $v;
        } else {
            $marital_statuses[] = $value;
        }
    }
    $marital_statuses = array_unique($marital_statuses);
    sort($marital_statuses);

    // Fetch job categories (taxonomy)
    $job_categories = [];
    $category_terms = get_terms([
        'taxonomy' => 'jobpost_category',
        'hide_empty' => false,
    ]);
    if (!is_wp_error($category_terms)) {
        foreach ($category_terms as $term) {
            $job_categories[] = $term->name;
        }
    }

    // Fetch job titles
    $job_titles = $wpdb->get_col("
        SELECT DISTINCT post_title 
        FROM {$wpdb->posts} 
        WHERE post_type = 'jobpost' AND post_status = 'publish'
        ORDER BY post_title ASC
    ");

    // Fetch job types (taxonomy)
    $job_types = [];
    $type_terms = get_terms([
        'taxonomy' => 'jobpost_job_type',
        'hide_empty' => false,
    ]);
    if (!is_wp_error($type_terms)) {
        foreach ($type_terms as $term) {
            $job_types[] = $term->name;
        }
    }

    // Fetch job locations (taxonomy)
    $job_locations = [];
    $location_terms = get_terms([
        'taxonomy' => 'jobpost_location',
        'hide_empty' => false,
    ]);
    if (!is_wp_error($location_terms)) {
        foreach ($location_terms as $term) {
            $job_locations[] = $term->name;
        }
    }

    ?>
    <div class="wrap">
        <h1 style="color:#21759b; font-weight:bold;">Custom Excel Export</h1>
        <form method="get" action="<?php echo admin_url('admin-post.php'); ?>" style="background:#f7f7f7; padding:20px; border-radius:8px;">
            <input type="hidden" name="action" value="cee_export_excel">

            <h2 style="color:#555;">Filter Applicants</h2>

            <p>
                <label for="date_from" style="margin-right:10px;">Date From:</label>
                <input type="date" name="date_from" id="date_from" style="margin-right:20px;">
                <label for="date_to" style="margin-right:10px;">Date To:</label>
                <input type="date" name="date_to" id="date_to">
            </p>

            <p>
                <label for="gender" style="margin-right:10px;">Gender:</label>
                <select name="gender" id="gender" style="margin-right:20px;">
                    <option value="">--All--</option>
                    <?php foreach ($genders as $g) echo '<option value="'.esc_attr($g).'">'.esc_html($g).'</option>'; ?>
                </select>

                <label for="marital_status" style="margin-right:10px;">Marital Status:</label>
                <select name="marital_status" id="marital_status" style="margin-right:20px;">
                    <option value="">--All--</option>
                    <?php foreach ($marital_statuses as $m) echo '<option value="'.esc_attr($m).'">'.esc_html($m).'</option>'; ?>
                </select>

                <label for="job_category" style="margin-right:10px;">Job Category:</label>
                <select name="job_category" id="job_category" style="margin-right:20px;">
                    <option value="">--All--</option>
                    <?php foreach ($job_categories as $c) echo '<option value="'.esc_attr($c).'">'.esc_html($c).'</option>'; ?>
                </select>

                <label for="job_title" style="margin-right:10px;">Job Title:</label>
                <select name="job_title" id="job_title" style="margin-right:20px;">
                    <option value="">--All--</option>
                    <?php foreach ($job_titles as $t) echo '<option value="'.esc_attr($t).'">'.esc_html($t).'</option>'; ?>
                </select>

                <label for="job_type" style="margin-right:10px;">Job Type:</label>
                <select name="job_type" id="job_type" style="margin-right:20px;">
                    <option value="">--All--</option>
                    <?php foreach ($job_types as $t) echo '<option value="'.esc_attr($t).'">'.esc_html($t).'</option>'; ?>
                </select>

                <label for="job_location" style="margin-right:10px;">Job Location:</label>
                <select name="job_location" id="job_location" style="margin-right:20px;">
                    <option value="">--All--</option>
                    <?php foreach ($job_locations as $l) echo '<option value="'.esc_attr($l).'">'.esc_html($l).'</option>'; ?>
                </select>
            </p>

            <p><input type="submit" class="button button-primary" value="Download Excel"></p>
        </form>
    </div>
    <?php
}

// Hook export
add_action('admin_post_cee_export_excel', 'cee_export_excel');

function cee_export_excel() {
    if (!current_user_can('manage_options')) wp_die('Unauthorized');

    global $wpdb;

    $date_from = $_GET['date_from'] ?? '';
    $date_to   = $_GET['date_to'] ?? '';
    $gender    = $_GET['gender'] ?? '';
    $marital_status = $_GET['marital_status'] ?? '';
    $job_category = $_GET['job_category'] ?? '';
    $job_title    = $_GET['job_title'] ?? '';
    $job_type     = $_GET['job_type'] ?? '';
    $job_location = $_GET['job_location'] ?? '';

    $args = [
        'post_type' => 'jobpost_applicants',
        'numberposts' => -1,
    ];

    // Date filter
    if ($date_from || $date_to) {
        $date_query = [];
        if ($date_from) $date_query['after'] = $date_from;
        if ($date_to) $date_query['before'] = $date_to;
        $date_query['inclusive'] = true;
        $args['date_query'] = [$date_query];
    }

    // Meta filters
    $meta_query = ['relation' => 'AND'];
    if ($gender) $meta_query[] = ['key' => 'jobapp_gender', 'value' => $gender, 'compare' => '='];
    if ($marital_status) $meta_query[] = ['key' => 'jobapp_marital_status', 'value' => $marital_status, 'compare' => '='];

    // Job title filter (match post by title and filter by stored job ID meta)
    if ($job_title) {
        $job_post = get_page_by_title($job_title, OBJECT, 'jobpost');
        if ($job_post) $meta_query[] = ['key' => 'jobapp_job_id', 'value' => $job_post->ID, 'compare' => '='];
    }

    if (count($meta_query) > 1) $args['meta_query'] = $meta_query;

    // Taxonomy filters
    $tax_query = ['relation' => 'AND'];
    if ($job_category) {
        $tax_query[] = [
            'taxonomy' => 'jobpost_category',
            'field' => 'name',
            'terms' => $job_category,
        ];
    }
    if ($job_type) {
        $tax_query[] = [
            'taxonomy' => 'jobpost_job_type',
            'field' => 'name',
            'terms' => $job_type,
        ];
    }
    if ($job_location) {
        $tax_query[] = [
            'taxonomy' => 'jobpost_location',
            'field' => 'name',
            'terms' => $job_location,
        ];
    }
    if (count($tax_query) > 1) $args['tax_query'] = $tax_query;

    $apps = get_posts($args);

    // --- Spreadsheet setup ---
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    $main_headers = [
        'S/N', 'Full Name', 'Gender', 'Age', 'Marital Status', 'Birth Place (Region)',
        'Birth Place (City/Town)', 'Current Address', 'Level of Education', 'Field of Study',
        'Year of Education', 'University/ College', 'CGPA'
    ];

    $col = 'A';
    foreach ($main_headers as $header) {
        $sheet->setCellValue($col . '1', $header);
        $sheet->mergeCells($col . '1:' . $col . '2');
        $sheet->getStyle($col . '1')->getAlignment()
            ->setHorizontal(Alignment::HORIZONTAL_CENTER)
            ->setVertical(Alignment::VERTICAL_CENTER);
        $sheet->getStyle($col . '1')->getFont()->setBold(true);
        $col++;
    }

    // Experience header
    $sheet->mergeCells('N1:R1');
    $sheet->setCellValue('N1', 'Three Years of Experience');
    $sheet->getStyle('N1')->getAlignment()
        ->setHorizontal(Alignment::HORIZONTAL_CENTER)
        ->setVertical(Alignment::VERTICAL_CENTER);
    $sheet->getStyle('N1')->getFont()->setBold(true);

    $exp_headers = ['Organization', 'Position', 'From–To (E.C)', 'Year of Exp.', 'Total Year of Exp.'];
    $sheet->fromArray($exp_headers, NULL, 'N2');
    $sheet->getStyle('N2:R2')->getAlignment()
        ->setHorizontal(Alignment::HORIZONTAL_CENTER)
        ->setVertical(Alignment::VERTICAL_CENTER);
    $sheet->getStyle('N2:R2')->getFont()->setBold(true);

    $other_headers = ['Current Gross Salary', 'Expected Gross Salary', 'Telephone'];
    $col = 'S';
    foreach ($other_headers as $header) {
        $sheet->setCellValue($col . '1', $header);
        $sheet->mergeCells($col . '1:' . $col . '2');
        $sheet->getStyle($col . '1')->getAlignment()
            ->setHorizontal(Alignment::HORIZONTAL_CENTER)
            ->setVertical(Alignment::VERTICAL_CENTER);
        $sheet->getStyle($col . '1')->getFont()->setBold(true);
        $col++;
    }

    // Fill data rows
    $row_num = 3;
    foreach ($apps as $i => $app) {
        $meta = get_post_meta($app->ID);

        $experiences = [
            [
                'company' => $meta['jobapp_current_company'][0] ?? '',
                'position' => $meta['jobapp_current_position'][0] ?? '',
                'from' => $meta['jobapp_experience_in_current_company_from'][0] ?? '',
                'to' => $meta['jobapp_experience_in_current_company_to'][0] ?? '',
                'years' => calculate_years($meta['jobapp_experience_in_current_company_from'][0] ?? '', $meta['jobapp_experience_in_current_company_to'][0] ?? ''),
                'total' => $meta['jobapp_total_year_of_experience'][0] ?? ''
            ],
            [
                'company' => $meta['jobapp_previous_company_1'][0] ?? '',
                'position' => $meta['jobapp_previous_position_in_company_1'][0] ?? '',
                'from' => $meta['jobapp_experience_in_previous_company_1_from'][0] ?? '',
                'to' => $meta['jobapp_experience_in_previous_company_1_to'][0] ?? '',
                'years' => calculate_years($meta['jobapp_experience_in_previous_company_1_from'][0] ?? '', $meta['jobapp_experience_in_previous_company_1_to'][0] ?? ''),
                'total' => ''
            ],
            [
                'company' => $meta['jobapp_prev_company2'][0] ?? '',
                'position' => $meta['jobapp_prev_position2'][0] ?? '',
                'from' => $meta['jobapp_prev_from2'][0] ?? '',
                'to' => $meta['jobapp_prev_to2'][0] ?? '',
                'years' => calculate_years($meta['jobapp_prev_from2'][0] ?? '', $meta['jobapp_prev_to2'][0] ?? ''),
                'total' => ''
            ]
        ];

        $experiences = array_filter($experiences, function ($exp) {
            return $exp['company'] || $exp['position'] || $exp['from'] || $exp['to'] || $exp['years'];
        });

        $exp_count = count($experiences);
        $start_row = $row_num;

        $main_cols = range('A', 'M');
        foreach ($main_cols as $mc) {
            $value = '';
            switch ($mc) {
                case 'A': $value = $i + 1; break;
                case 'B': $value = $meta['jobapp_name'][0] ?? ''; break;
                case 'C': $value = $meta['jobapp_gender'][0] ?? ''; break;
                case 'D': $value = $meta['jobapp_age'][0] ?? ''; break;
                case 'E': $value = $meta['jobapp_marital_status'][0] ?? ''; break;
                case 'F': $value = $meta['jobapp_birth_place_state'][0] ?? ''; break; // fixed key
                case 'G': $value = $meta['jobapp_birth_place_city'][0] ?? ''; break;
                case 'H': $value = $meta['jobapp_current_address'][0] ?? ''; break;
                case 'I': $value = $meta['jobapp_educational_level'][0] ?? ''; break;
                case 'J': $value = $meta['jobapp_field_of_study'][0] ?? ''; break;
                case 'K': $value = $meta['jobapp_year_of_graduation'][0] ?? ''; break;
                case 'L': $value = $meta['jobapp_university_college'][0] ?? ''; break;
                case 'M': $value = $meta['jobapp_cgpa'][0] ?? ''; break;
            }
            $sheet->setCellValue($mc . $row_num, $value);
            if ($exp_count > 1) $sheet->mergeCells($mc . $row_num . ':' . $mc . ($row_num + $exp_count - 1));
            $sheet->getStyle($mc . $row_num)->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
        }

        foreach ($experiences as $exp) {
            $sheet->setCellValue('N' . $row_num, $exp['company']);
            $sheet->setCellValue('O' . $row_num, $exp['position']);
            $sheet->setCellValue('P' . $row_num, $exp['from'] . '–' . $exp['to']);
            $sheet->setCellValue('Q' . $row_num, $exp['years']);
            $sheet->setCellValue('R' . $row_num, $exp['total']);
            $row_num++;
        }

        $salary_cols = ['S', 'T', 'U'];
        foreach ($salary_cols as $sc) {
            $value = '';
            switch ($sc) {
                case 'S': $value = $meta['jobapp_current_gross_salary'][0] ?? ''; break;
                case 'T': $value = $meta['jobapp_expected_gross_salary'][0] ?? ''; break;
                case 'U': $value = $meta['jobapp_phone'][0] ?? ''; break;
            }
            $sheet->setCellValue($sc . $start_row, $value);
            if ($exp_count > 1) $sheet->mergeCells($sc . $start_row . ':' . $sc . ($row_num - 1));
            $sheet->getStyle($sc . $start_row)->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
        }

        for ($r = $start_row; $r < $row_num; $r++) {
            $fillColor = ($r % 2 == 0) ? 'FFEFEFEF' : 'FFFFFFFF';
            $sheet->getStyle('A' . $r . ':U' . $r)
                ->getFill()->setFillType(Fill::FILL_SOLID)
                ->getStartColor()->setARGB($fillColor);
        }
    }

    // Download
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="applicants.xlsx"');
    header('Cache-Control: max-age=0');

    $writer = new Xlsx($spreadsheet);
    $writer->save('php://output');
    exit;
}

// Utility
function calculate_years($from, $to) {
    if (!$from || !$to) return '';
    $from = strtotime($from);
    $to = strtotime($to);
    return $from && $to ? round(abs($to - $from) / (365*24*60*60), 2) : '';
}
