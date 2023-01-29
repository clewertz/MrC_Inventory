<!DOCTYPE html>
<html>
<head>
	<title>Laptop Request From</title>
	<script type="text/javascript" src="inc/js/bootstrap.min.js"></script>
	<link rel="stylesheet" type="text/css" href="inc/laptop.css">
	<link rel="stylesheet" type="text/css" href="inc/css/bootstrap.min.css">
</head>
<body>

<div class="container">

    <form class="well form-horizontal" action=" " method="post"  id="contact_form">
<fieldset>

<!-- Form Name -->
<legend>Mr. Cooper Laptop Request Form</legend>

<!-- Text input FIRST NAME-->

<div class="form-group">
  <label class="col-md-4 control-label">First Name</label>  
  <div class="col-md-4 inputGroupContainer">
  <div class="input-group">
  <span class="input-group-addon"><i class="glyphicon glyphicon-user"></i></span>
  <input  name="first_name" placeholder="First Name" class="form-control"  type="text">
    </div>
  </div>
</div>

<!-- Text input LAST NAME-->

<div class="form-group">
  <label class="col-md-4 control-label" >Last Name</label> 
    <div class="col-md-4 inputGroupContainer">
    <div class="input-group">
  <span class="input-group-addon"><i class="glyphicon glyphicon-user"></i></span>
  <input name="last_name" placeholder="Last Name" class="form-control"  type="text">
    </div>
  </div>
</div>

<!-- Text input CUBE/OFFICE NUMBER-->
       <div class="form-group">
  <label class="col-md-4 control-label">Cube/Office Number</label>  
    <div class="col-md-4 inputGroupContainer">
    <div class="input-group">
        <span class="input-group-addon"><i class="glyphicon glyphicon-envelope"></i></span>
  <input name="email" placeholder="Cube/Office Number" class="form-control"  type="text">
    </div>
  </div>
</div>

<!-- Checkbox for Laptop -->
 <div class="form-group">
    <label class="col-md-4 control-label">Please Select the Type of Laptop:</label>
 <div class="col-md-4">
 	<div class="radio">
    	<label>
        	<input type="radio" name="laptop-type" value="thin" /> Citrix Laptop
    	</label> 
 	</div>
    <div class="radio">
    <label>
        <input type="radio" name="laptop-type" value="thick" /> Windows Laptop*
    </label>
   </div>
  </div>
 </div>

<div>
	<h5><strong>*Windows based laptops require an Archer exception.</strong></h5>
	<h5>If a Windows laptop was selected, please provide the Archer Exception number for this request.</h5>
</div>

<!-- Text input ARCHER EXCEPTION-->
       
<div class="form-group">
  <label class="col-md-4 control-label">Archer Exception</label>  
    <div class="col-md-4 inputGroupContainer">
    <div class="input-group">
        <span class="input-group-addon"><i class="glyphicon glyphicon-ok-circle"></i></span>
  <input name="phone" placeholder="EXC-####" class="form-control" type="text">
    </div>
  </div>
</div>

<!-- Text area -->
<div>
	<h5>Please provide a detailed business justification for the purchase of a laptop.</h5>	
	<h5>This justification will be reviewed by your management and the IT department.</h5>
</div>

  
<div class="form-group">
  <label class="col-md-4 control-label"></label>
    <div class="col-md-4 inputGroupContainer">
    <div class="input-group">
        <span class="input-group-addon"><i class="glyphicon glyphicon-pencil"></i></span>
        	<textarea class="form-control" name="comment" placeholder="Enter Detailed Business Justification"></textarea>
  </div>
  </div>
</div>

<!-- Checkboxes for Accessories -->
 <div class="form-group">
              <label class="col-md-4 control-label">Addtional Accesories Needed</label>
                        <div class="col-md-4">
                            <div class="checkbox">
                                <label>
                                    <input type="checkbox" name="dock" value="yes" /> Docking Station
                                </label>
                            </div>
                            <div class="checkbox">
                                <label>
                                    <input type="checkbox" name="lock cable" value="yes" /> Security Lock Cable
                                </label>
                            <div class="checkbox">
                                <label>
                                    <input type="checkbox" name="bag" value="yes" /> Laptop Bag
                                </label>
                            </div>
                            <div class="checkbox">
                                <label>
                                    <input type="checkbox" name="other" value="yes" /> Other (Please detail below)
                                </label>
                            </div>    
                            </div>
                        </div>
                    </div>

<!-- Text area -->
  
<div class="form-group">
  <label class="col-md-4 control-label"></label>
    <div class="col-md-4 inputGroupContainer">
    <div class="input-group">
        <span class="input-group-addon"><i class="glyphicon glyphicon-pencil"></i></span>
        	<textarea class="form-control" name="comment" placeholder="Enter Detailed Business Justification"></textarea>
  </div>
  </div>
</div>

<!-- Success message -->
<div class="alert alert-success" role="alert" id="success_message">Success <i class="glyphicon glyphicon-thumbs-up"></i> Thank you!  You will ne notified once your laptop is ready for deployment.</div>

<!-- Button -->
<div class="form-group">
  <label class="col-md-4 control-label"></label>
  <div class="col-md-4">
    <button type="submit" class="btn btn-warning" >Submit <span class="glyphicon glyphicon-send"></span></button>
  </div>
</div>

</fieldset>
</form>
</div>
    </div><!-- /.container -->

<!-- Text input-->
      
<!-- <div class="form-group">
  <label class="col-md-4 control-label">Address</label>  
    <div class="col-md-4 inputGroupContainer">
    <div class="input-group">
        <span class="input-group-addon"><i class="glyphicon glyphicon-home"></i></span>
  <input name="address" placeholder="Address" class="form-control" type="text">
    </div>
  </div>
</div> -->

<!-- Text input-->
 
<!-- <div class="form-group">
  <label class="col-md-4 control-label">City</label>  
    <div class="col-md-4 inputGroupContainer">
    <div class="input-group">
        <span class="input-group-addon"><i class="glyphicon glyphicon-home"></i></span>
  <input name="city" placeholder="city" class="form-control"  type="text">
    </div>
  </div>
</div> -->

<!-- Select Basic -->
   
<!-- <div class="form-group"> 
  <label class="col-md-4 control-label">State</label>
    <div class="col-md-4 selectContainer">
    <div class="input-group">
        <span class="input-group-addon"><i class="glyphicon glyphicon-list"></i></span>
    <select name="state" class="form-control selectpicker" >
      <option value=" " >Please select your state</option>
      <option>Alabama</option>
      <option>Alaska</option>
      <option >Arizona</option>
      <option >Arkansas</option>
      <option >California</option>
      <option >Colorado</option>
      <option >Connecticut</option>
      <option >Delaware</option>
      <option >District of Columbia</option>
      <option> Florida</option>
      <option >Georgia</option>
      <option >Hawaii</option>
      <option >daho</option>
      <option >Illinois</option>
      <option >Indiana</option>
      <option >Iowa</option>
      <option> Kansas</option>
      <option >Kentucky</option>
      <option >Louisiana</option>
      <option>Maine</option>
      <option >Maryland</option>
      <option> Mass</option>
      <option >Michigan</option>
      <option >Minnesota</option>
      <option>Mississippi</option>
      <option>Missouri</option>
      <option>Montana</option>
      <option>Nebraska</option>
      <option>Nevada</option>
      <option>New Hampshire</option>
      <option>New Jersey</option>
      <option>New Mexico</option>
      <option>New York</option>
      <option>North Carolina</option>
      <option>North Dakota</option>
      <option>Ohio</option>
      <option>Oklahoma</option>
      <option>Oregon</option>
      <option>Pennsylvania</option>
      <option>Rhode Island</option>
      <option>South Carolina</option>
      <option>South Dakota</option>
      <option>Tennessee</option>
      <option>Texas</option>
      <option> Uttah</option>
      <option>Vermont</option>
      <option>Virginia</option>
      <option >Washington</option>
      <option >West Virginia</option>
      <option>Wisconsin</option>
      <option >Wyoming</option>
    </select>
  </div>
</div>
</div> -->

<!-- Text input-->

<!-- <div class="form-group">
  <label class="col-md-4 control-label">Zip Code</label>  
    <div class="col-md-4 inputGroupContainer">
    <div class="input-group">
        <span class="input-group-addon"><i class="glyphicon glyphicon-home"></i></span>
  <input name="zip" placeholder="Zip Code" class="form-control"  type="text">
    </div>
</div>
</div> -->

<!-- Text input-->
<!-- <div class="form-group">
  <label class="col-md-4 control-label">Website or domain name</label>  
   <div class="col-md-4 inputGroupContainer">
    <div class="input-group">
        <span class="input-group-addon"><i class="glyphicon glyphicon-globe"></i></span>
  <input name="website" placeholder="Website or domain name" class="form-control" type="text">
    </div>
  </div>
</div> -->


	<script type="text/javascript" src="inc/laptop.js"></script>
</body>
</html>