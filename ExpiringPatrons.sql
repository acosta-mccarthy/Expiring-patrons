SELECT
CONCAT (patron_record_fullname.first_name, ' ', patron_record_fullname.middle_name, ' ', patron_record_fullname.last_name) AS "PATRON NAME",
barcode AS "PATRON BARCODE",
TO_CHAR(expiration_date_gmt, 'MM/DD/YYYY') AS "EXPIRATION DATE",
home_library_code AS "HOME LIBRARY",
field_content AS "EMAIL ADDRESS"

FROM
sierra_view.patron_view
JOIN sierra_view.patron_record_fullname
ON patron_record_fullname.patron_record_id = patron_view.id
JOIN sierra_view.varfield_view
ON varfield_view.record_id = patron_view.id

WHERE
varfield_type_code = 'z' AND
expiration_date_gmt >= DATE_TRUNC('month', now()) + interval '1 month' AND
expiration_date_gmt < DATE_TRUNC('month', now()) + interval '2 months'
--Finds all patrons with an email address on file that have expiration dates after the current month, and before 2 months from now

ORDER BY "HOME LIBRARY"