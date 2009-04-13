<?php head(array('title' => 'Spreadsheet', 'body_class' => 'speadsheet-plugin')); ?>
<h1>Spreadsheet Status</h1>

<div id="primary">
	<?php echo flash(); ?>
	<div id="speadsheet">
		<p>Your most recent export is listed first. Large exports may take some time.  If the export you need shows "in-progress", refresh this page periodically until you see "complete". Then you may download your spreadsheet. Your previous exports, if any, are also available to re-download here.</p>
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
					<td><?php echo $e->file_name?></td>
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