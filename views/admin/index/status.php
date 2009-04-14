<?php head(array('title' => 'Spreadsheet', 'body_class' => 'speadsheet-plugin')); ?>
<h1>Spreadsheet Status</h1>

<div id="primary">
	<?php echo flash(); ?>
	<div id="speadsheet">
		<p>Please note:
			<ul style="list-style-type:disc;margin-left:20px">
				<li>Only <strong>your</strong> exports are listed here.</li>
				<li>Exports are listed most recent first along with the search terms used to produce the export.</li> 
				<li>Large exports may take some time.  If the export you need shows "in-progress", refresh this page periodically until you see "complete". Then you may download your spreadsheet by clicking on the "download" link.</li>
				<li>Your previous exports, if any, are also available to re-download here.</li>
			</ul>
		</p>
		<table class="simple">
			<tr>
				<th>Date</th>
				<th>File</th>
				<th>Status</th>
				<th></th>
			</tr>
			<?php foreach ($this->exports as $e) { ?>
				<tr>
					<td><?php echo $e->added ?></td>
					<td>
						<?php echo $e->file_name?><br />
						<?php
							$terms = unserialize($e->terms);
							$out = "";
							foreach ($terms as $k => $v) {
								if ($k == 'advanced') {
									foreach($v as $advanced_term) {
										 $out .= $advanced_term['element_id'] . ' ' . $advanced_term['type'] . ' ' . $advanced_term['terms'] . ' ';
									}
								} else if ($k == 'submit_search' || !$v) {
									continue;
								} else {
									$out .= " <strong>{$k}</strong>: {$v};";
								}
							}
							echo $out;
						?>
					</td>
					<td>
						<?php
						switch ($e->status) {
							case SPREADSHEET_STATUS_INIT:
								echo 'in-progress';
								break;
							case SPREADSHEET_STATUS_COMPLETE:
							  echo 'complete';
								break;
							case SPREADSHEET_STATUS_ERROR:
								echo 'error';
						} 						
						?>
					</td>
					<td>
						<?php if ($e->status == SPREADSHEET_STATUS_COMPLETE) { ?>
							<a href="<?php echo uri('spreadsheet/download/' . $e->id) ?>">Download</a>
						<?php } ?>
					</td>
				</tr>
			<?php } ?>
		</table>
	</div>
	
</div>

<?php foot(); ?>