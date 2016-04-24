from django.shortcuts import render
from django.http import HttpResponse, HttpResponseRedirect
from django.core.urlresolvers import reverse
from tutorial.authhelper import get_signin_url, get_token_from_code, get_user_email_from_id_token
from tutorial.outlookservice import get_my_events, get_my_messages, get_my_contacts
# Create your views here.


def home(request):
    redirect_uri = request.build_absolute_uri(reverse('tutorial:gettoken'))
    sign_in_url = get_signin_url(redirect_uri)

    return HttpResponse('<a href="' + sign_in_url +'">Click here to sign in and view your mail</a>')


def gettoken(request):
    auth_code = request.GET['code']
    redirect_uri = request.build_absolute_uri(reverse('tutorial:gettoken'))
    token = get_token_from_code(auth_code, redirect_uri)
    access_token = token['access_token']

    # Save the token in the session
    request.session['access_token'] = access_token
    return HttpResponseRedirect(reverse('tutorial:mail'))


def mail(request):
    access_token = request.session['access_token']
    # If there is no token in the session, redirect to home
    if not access_token:
        return HttpResponseRedirect(reverse('tutorial:home'))
    else:
        messages = get_my_messages(access_token)
        return HttpResponse('Messages: {0}'.format(messages))


def events(request):
    access_token = request.session['access_token']
    # If there is no token in the session, redirect to home
    if not access_token:
        return HttpResponseRedirect(reverse('tutorial:home'))
    else:
        events = get_my_events(access_token)
        return HttpResponse('Events: {0}'.format(events))


def contacts(request):
    access_token = request.session['access_token']
    # If there is no token in the session, redirect to home
    if not access_token:
        return HttpResponseRedirect(reverse('tutorial:home'))
    else:
        contacts = get_my_contacts(access_token)
        return HttpResponse('Contacts: {0}'.format(contacts))
