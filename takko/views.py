from django.shortcuts import render

# Create your views here.
from django.http import HttpResponse
from django.http import HttpResponseRedirect
from django.utils.encoding import smart_str


from .models import UploadFileForm
from .models import TakkoOrder
from .models import TakkoInvoice
from .models import NewTakkoOrder

import os


def index(request):
    return HttpResponse("Hello, world. You're at the takko index.")


fileDir = 'takkoUploadedFile'


def upload_file(request):
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            save_uploaded_file(request.FILES['file'])
            #fileName = combineOrders.combineOrders(fileDir)
            takko_order = NewTakkoOrder(fileDir)
            fileName = takko_order.save_to_excel()
            return download_file(fileName)
            #return HttpResponseRedirect('/success/url/')
    else:
        form = UploadFileForm()
    return render(request, 'upload.html', {'form': form})


def save_uploaded_file(f):
    with open(fileDir, 'wb+') as destination:
        for chunk in f.chunks():
            destination.write(chunk)


def download_file(fileDir):
    with open(fileDir, 'rb') as fh:
        response = HttpResponse(fh.read(), content_type="application/force-download")
        #response['Content-Disposition'] = 'inline; filename=' + os.path.basename(fileDir)
        #return response

        response['Content-Disposition'] = 'attachment; filename=%s' % smart_str(fileDir)
        response['X-Sendfile'] = smart_str(fileDir)
        # It's usually a good idea to set the 'Content-Length' header too.
        # You can also set any other required headers: Cache-Control, etc.
        return response


def invoice_test(request):
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            save_uploaded_file(request.FILES['file'])
            #fileName = matchInvoiceNumbers.matchInvoiceNumbers(fileDir)
            takko_invoice = TakkoInvoice(fileDir)
            fileName = takko_invoice.save_to_excel()
            return download_file(fileName)
    else:
        form = UploadFileForm()
    return render(request, 'invoice_test.html', {'form': form})
